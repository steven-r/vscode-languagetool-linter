/****
 *    Copyright 2019 David L. Day
 *
 *   Licensed under the Apache License, Version 2.0 (the "License");
 *   you may not use this file except in compliance with the License.
 *   You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 *   Unless required by applicable law or agreed to in writing, software
 *   distributed under the License is distributed on an "AS IS" BASIS,
 *   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *   See the License for the specific language governing permissions and
 *   limitations under the License.
 */

import * as RehypeBuilder from "annotatedtext-rehype";
import * as RemarkBuilder from "./annotation-builder";
import * as Fetch from "node-fetch";
import fs from "fs";
import os from "os";
import path from "path";

import {
  CancellationToken,
  CodeAction,
  CodeActionContext,
  CodeActionKind,
  CodeActionProvider,
  Diagnostic,
  DiagnosticCollection,
  DiagnosticSeverity,
  languages,
  Position,
  Range,
  TextDocument,
  Uri,
  window,
  workspace,
  WorkspaceEdit,
} from "vscode";
import * as Constants from "./Constants";
import { ConfigurationManager } from "./ConfigurationManager";
import { FormattingProviderDashes } from "./FormattingProviderDashes";
import { FormattingProviderEllipses } from "./FormattingProviderEllipses";
import { FormattingProviderQuotes } from "./FormattingProviderQuotes";
import {
  ILanguageToolMatch,
  ILanguageToolReplacement,
  ILanguageToolResponse,
} from "./Interfaces";
import { IAnnotatedtext, IAnnotation } from "annotatedtext";
import { findUpSync } from "find-up";

class LTDiagnostic extends Diagnostic {
  match?: ILanguageToolMatch;
}

export class Linter implements CodeActionProvider {
  // Is the rule a Spelling rule?
  // See: https://forum.languagetool.org/t/identify-spelling-rules/4775/3
  public static isSpellingRule(ruleId: string): boolean {
    return (
      ruleId.indexOf("MORFOLOGIK_RULE") !== -1 ||
      ruleId.indexOf("SPELLER_RULE") !== -1 ||
      ruleId.indexOf("HUNSPELL_NO_SUGGEST_RULE") !== -1 ||
      ruleId.indexOf("HUNSPELL_RULE") !== -1 ||
      ruleId.indexOf("GERMAN_SPELLER_RULE") !== -1 ||
      ruleId.indexOf("FR_SPELLING_RULE") !== -1
    );
  }

  public diagnosticCollection: DiagnosticCollection;
  public remarkBuilderOptions: RemarkBuilder.IOptions = RemarkBuilder.defaults;
  public rehypeBuilderOptions: RehypeBuilder.IOptions = RehypeBuilder.defaults;

  private readonly configManager: ConfigurationManager;
  private timeoutMap: Map<string, NodeJS.Timeout>;
  private enabledDisabledRules: Record<string, boolean> = {};
  private enabledDisabledRulesPerLine: Record<number, Record<string, boolean>> =
    {};
  private globalEnabledDisabledRules: Record<string, boolean> = {};

  // Regular expression for matching the start of inline disable/enable comments (from https://github.com/DavidAnson/markdownlint/blob/main/helpers/helpers.js)
  private inlineCommentStartRe =
    /(<!--\s*languagetool-(disable-file|enable-file|disable-line|disable-next-line|configure-file))(?:\s|-->)/gi;
  private configFileStartRe =
    /^\s*(languagetool-)?(?<command>disable|enable)(?<parameter>(\s+([A-Z_0-9]+)(\([^)]+\))?)+)/gi;
  private inlineCommentRuleRe = /\s*([A-Z_0-9]+)(\((.+?)\))?/gi;

  constructor(configManager: ConfigurationManager) {
    this.configManager = configManager;
    this.timeoutMap = new Map<string, NodeJS.Timeout>();
    this.diagnosticCollection = languages.createDiagnosticCollection(
      Constants.EXTENSION_DISPLAY_NAME,
    );
    this.remarkBuilderOptions.interpretmarkup = this.customMarkdownInterpreter;
    this.readConfigurationFiles(configManager.getConfigurationFiles());
    this.globalEnabledDisabledRules = this.enabledDisabledRules;
  }

  // Provide CodeActions for thw given Document and Range
  public provideCodeActions(
    document: TextDocument,
    _range: Range,
    context: CodeActionContext,
    _token: CancellationToken,
  ): CodeAction[] {
    const diagnostics = context.diagnostics || [];
    const actions: CodeAction[] = [];
    diagnostics
      .filter(
        (diagnostic) =>
          diagnostic.source === Constants.EXTENSION_DIAGNOSTIC_SOURCE,
      )
      .forEach((diagnostic) => {
        const match:
          | ILanguageToolMatch
          | undefined = (diagnostic as LTDiagnostic).match;
        if (match && Linter.isSpellingRule(match.rule.id)) {
          const spellingActions: CodeAction[] = this.getSpellingRuleActions(
            document,
            diagnostic,
          );
          spellingActions.forEach((action) => {
            actions.push(action);
          });
        } else if (match) {
          const word: string = document.getText(diagnostic.range);
          this.addLocalFileIgnore(document, word, match, diagnostic).forEach(
            (action: CodeAction) => {
              actions.push(action);
            },
          );
          this.getRuleActions(document, diagnostic).forEach((action) => {
            actions.push(action);
          });
        }
      });
    return actions;
  }

  // Remove diagnostics for a Document URI
  public clearDiagnostics(uri: Uri): void {
    this.diagnosticCollection.delete(uri);
  }

  // Request a lint for a document
  public requestLint(
    document: TextDocument,
    timeoutDuration: number = Constants.EXTENSION_TIMEOUT_MS,
  ): void {
    if (this.configManager.isSupportedDocument(document)) {
      this.cancelLint(document);
      const uriString = document.uri.toString();
      const timeout = setTimeout(() => {
        this.lintDocument(document);
      }, timeoutDuration);
      this.timeoutMap.set(uriString, timeout);
    }
  }
  // Force request a lint for a document as plain text regardless of language id
  public requestLintAsPlainText(
    document: TextDocument,
    timeoutDuration: number = Constants.EXTENSION_TIMEOUT_MS,
  ): void {
    this.cancelLint(document);
    const uriString = document.uri.toString();
    const timeout = setTimeout(() => {
      this.lintDocumentAsPlainText(document);
    }, timeoutDuration);
    this.timeoutMap.set(uriString, timeout);
  }

  // Cancel lint
  public cancelLint(document: TextDocument): void {
    const uriString: string = document.uri.toString();
    if (this.timeoutMap.has(uriString)) {
      if (this.timeoutMap.has(uriString)) {
        const timeout: NodeJS.Timeout = this.timeoutMap.get(
          uriString,
        ) as NodeJS.Timeout;
        clearTimeout(timeout);
        this.timeoutMap.delete(uriString);
      }
    }
  }

  // Build annotatedtext from Markdown
  public buildAnnotatedMarkdown(text: string): IAnnotatedtext {
    return RemarkBuilder.build(text, this.remarkBuilderOptions);
  }

  // Build annotatedtext from HTML
  public buildAnnotatedHTML(text: string): IAnnotatedtext {
    return RehypeBuilder.build(text, this.rehypeBuilderOptions);
  }

  // Build annotatedtext from PLAINTEXT
  public buildAnnotatedPlaintext(plainText: string): IAnnotatedtext {
    const textAnnotation: IAnnotation = {
      text: plainText,
      offset: {
        start: 0,
        end: plainText.length,
      },
    };
    return { annotation: [textAnnotation] };
  }

  // Abstract annotated text builder
  public buildAnnotatedtext(document: TextDocument): IAnnotatedtext {
    let annotatedtext: IAnnotatedtext = { annotation: [] };
    switch (document.languageId) {
      case Constants.LANGUAGE_ID_MARKDOWN:
        annotatedtext = this.buildAnnotatedMarkdown(document.getText());
        break;
      case Constants.LANGUAGE_ID_MDX:
        annotatedtext = this.buildAnnotatedMarkdown(document.getText());
        break;
      case Constants.LANGUAGE_ID_HTML:
        annotatedtext = this.buildAnnotatedHTML(document.getText());
        break;
      default:
        annotatedtext = this.buildAnnotatedPlaintext(document.getText());
        break;
    }
    return annotatedtext;
  }

  // Perform Lint on Document
  public lintDocument(document: TextDocument): void {
    if (this.configManager.isSupportedDocument(document)) {
      if (document.languageId === Constants.LANGUAGE_ID_MARKDOWN) {
        const annotatedMarkdown: string = JSON.stringify(
          this.buildAnnotatedMarkdown(document.getText()),
        );
        this.lintAnnotatedText(document, annotatedMarkdown);
      } else if (document.languageId === Constants.LANGUAGE_ID_HTML) {
        const annotatedHTML: string = JSON.stringify(
          this.buildAnnotatedHTML(document.getText()),
        );
        this.lintAnnotatedText(document, annotatedHTML);
      } else {
        this.lintDocumentAsPlainText(document);
      }
    }
  }

  // Perform Lint on Document As Plain Text
  public lintDocumentAsPlainText(document: TextDocument): void {
    const annotatedPlaintext: string = JSON.stringify(
      this.buildAnnotatedPlaintext(document.getText()),
    );
    this.lintAnnotatedText(document, annotatedPlaintext);
  }

  // Lint Annotated Text
  public lintAnnotatedText(
    document: TextDocument,
    annotatedText: string,
  ): void {
    const ltPostDataDict: Record<string, string> = this.getPostDataTemplate();
    ltPostDataDict.data = annotatedText;
    this.callLanguageTool(document, ltPostDataDict);
  }

  // Apply smart formatting to annotated text.
  public smartFormatAnnotatedtext(annotatedtext: IAnnotatedtext): string {
    let newText = "";
    // Only run substitutions on text annotations.
    annotatedtext.annotation.forEach((annotation) => {
      if (annotation.text) {
        newText += annotation.text
          // Open Double Quotes
          .replace(/"(?=[\w'‘])/g, FormattingProviderQuotes.startDoubleQuote)
          // Close Double Quotes
          .replace(
            /([\w.!?%,'’])"/g,
            "$1" + FormattingProviderQuotes.endDoubleQuote,
          )
          // Remaining Double Quotes
          .replace(/"/, FormattingProviderQuotes.endDoubleQuote)
          // Open Single Quotes
          .replace(
            /(\W)'(?=[\w"“])/g,
            "$1" + FormattingProviderQuotes.startSingleQuote,
          )
          // Closing Single Quotes
          .replace(
            /([\w.!?%,"”])'/g,
            "$1" + FormattingProviderQuotes.endSingleQuote,
          )
          // Remaining Single Quotes
          .replace(/'/, FormattingProviderQuotes.endSingleQuote)
          .replace(/([\w])---(?=[\w])/g, "$1" + FormattingProviderDashes.emDash)
          .replace(/([\w])--(?=[\w])/g, "$1" + FormattingProviderDashes.enDash)
          .replace(/\.\.\./g, FormattingProviderEllipses.ellipses);
      } else if (annotation.markup) {
        newText += annotation.markup;
      }
    });
    return newText;
  }

  // Private instance methods

  // Custom markdown interpretation
  private customMarkdownInterpreter(text: string): string {
    // Default of preserve line breaks
    let interpretation = "\n".repeat((text.match(/\n/g) || []).length);
    if (text.match(/^(?!\s*`{3})\s*`{1,2}/)) {
      // Treat inline code as redacted text
      interpretation = "`" + "#".repeat(text.length - 2) + "`";
    } else if (text.match(/#\s+$/)) {
      // Preserve Headers
      interpretation += "# ";
    } else if (text.match(/\*\s+$/)) {
      // Preserve bullets without leading spaces
      interpretation += "* ";
    } else if (text.match(/\d+\.\s+$/)) {
      // Treat as bullets without leading spaces
      interpretation += "** ";
    }
    return interpretation;
  }

  // Set ltPostDataTemplate from Configuration
  private getPostDataTemplate(): Record<string, string> {
    const ltPostDataTemplate: Record<string, string> = {};
    this.configManager.getServiceParameters().forEach((value, key) => {
      ltPostDataTemplate[key] = value;
    });
    return ltPostDataTemplate;
  }

  // Call to LanguageTool Service
  private callLanguageTool(
    document: TextDocument,
    ltPostDataDict: Record<string, string>,
  ): void {
    const url = this.configManager.getUrl();
    if (url) {
      const formBody = Object.keys(ltPostDataDict)
        .map(
          (key: string) =>
            encodeURIComponent(key) +
            "=" +
            encodeURIComponent(ltPostDataDict[key]),
        )
        .join("&");

      const options: Fetch.RequestInit = {
        body: formBody,
        headers: {
          "Accepts": "application/json",
          "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        },
        method: "POST",
      };
      Fetch.default(url, options)
        .then((res) => res.json())
        .then((json) => this.suggest(document, json))
        .catch((err) => {
          Constants.EXTENSION_OUTPUT_CHANNEL.appendLine(
            "Error connecting to " + url,
          );
          Constants.EXTENSION_OUTPUT_CHANNEL.appendLine(err);
        });
    } else {
      Constants.EXTENSION_OUTPUT_CHANNEL.appendLine(
        "No LanguageTool URL provided. Please check your settings and try again.",
      );
      Constants.EXTENSION_OUTPUT_CHANNEL.show(true);
    }
  }

  // Convert LanguageTool Suggestions into QuickFix CodeActions
  private suggest(
    document: TextDocument,
    response: ILanguageToolResponse,
  ): void {
    const matches = response.matches;
    const diagnostics: LTDiagnostic[] = [];
    this.buildRuleList(document);
    matches.forEach((match: ILanguageToolMatch) => {
      const start: Position = document.positionAt(match.offset);
      const end: Position = document.positionAt(match.offset + match.length);
      const diagnosticSeverity: DiagnosticSeverity =
        this.configManager.getDiagnosticSeverity();
      const diagnosticRange: Range = new Range(start, end);
      const diagnosticMessage: string = match.message;
      const diagnostic: LTDiagnostic = new LTDiagnostic(
        diagnosticRange,
        diagnosticMessage,
        diagnosticSeverity,
      );
      diagnostic.source = Constants.EXTENSION_DIAGNOSTIC_SOURCE;
      diagnostic.match = match;
      if (Linter.isSpellingRule(match.rule.id)) {
        if (!this.configManager.isHideRuleIds()) {
          diagnostic.code = match.rule.id;
        }
      } else {
        diagnostic.code = {
          target: this.configManager.getRuleUrl(match.rule.id),
          value: this.configManager.isHideRuleIds()
            ? Constants.SERVICE_RULE_URL_GENERIC_LABEL
            : match.rule.id,
        };
      }
      diagnostics.push(diagnostic);
      const word = document.getText(diagnostic.range);
      if (
        (Linter.isSpellingRule(match.rule.id) &&
          this.configManager.isIgnoredWord(word) &&
          this.configManager.showIgnoredWordHints()) ||
        this.checkIfRuleIsIgnored(match.rule.id, word, start)
      ) {
        diagnostic.severity = DiagnosticSeverity.Hint;
      }
    });
    this.diagnosticCollection.set(document.uri, diagnostics);
  }

  // Get CodeActions for Spelling Rules
  private getSpellingRuleActions(
    document: TextDocument,
    diagnostic: LTDiagnostic,
  ): CodeAction[] {
    const actions: CodeAction[] = [];
    const match: ILanguageToolMatch | undefined = diagnostic.match;
    const word: string = document.getText(diagnostic.range);
    if (this.configManager.isIgnoredWord(word)) {
      this.handleAlreadyIgnoredWord(word, diagnostic).forEach(
        (action: CodeAction) => actions.push(action),
      );
    } else {
      const usrIgnoreActionTitle: string = "Always ignore '" + word + "'";
      const usrIgnoreAction: CodeAction = new CodeAction(
        usrIgnoreActionTitle,
        CodeActionKind.QuickFix,
      );
      usrIgnoreAction.command = {
        arguments: [word],
        command: "languagetoolLinter.ignoreWordGlobally",
        title: usrIgnoreActionTitle,
      };
      usrIgnoreAction.diagnostics = [];
      usrIgnoreAction.diagnostics.push(diagnostic);
      actions.push(usrIgnoreAction);
      if (workspace !== undefined) {
        const wsIgnoreActionTitle: string =
          "Ignore '" + word + "' in Workspace";
        const wsIgnoreAction: CodeAction = new CodeAction(
          wsIgnoreActionTitle,
          CodeActionKind.QuickFix,
        );
        wsIgnoreAction.command = {
          arguments: [word],
          command: "languagetoolLinter.ignoreWordInWorkspace",
          title: wsIgnoreActionTitle,
        };
        wsIgnoreAction.diagnostics = [];
        wsIgnoreAction.diagnostics.push(diagnostic);
        actions.push(wsIgnoreAction);
      }
      if (match) {
        this.addLocalFileIgnore(document, word, match, diagnostic).forEach(
          (action: CodeAction) => {
            actions.push(action);
          },
        );
        this.getReplacementActions(
          document,
          diagnostic,
          match.replacements,
        ).forEach((action: CodeAction) => {
          actions.push(action);
        });
      }
    }
    return actions;
  }

  private addLocalFileIgnore(
    document: TextDocument,
    word: string,
    match: ILanguageToolMatch,
    diagnostic: LTDiagnostic,
  ): CodeAction[] {
    const actions: CodeAction[] = [];
    if (
      document.languageId == Constants.LANGUAGE_ID_MARKDOWN ||
      document.languageId == Constants.LANGUAGE_ID_HTML
    ) {
      const title: string = "Ignore '" + word + "' at current occourence";
      const wsIgnoreAction: CodeAction = new CodeAction(
        title,
        CodeActionKind.QuickFix,
      );
      wsIgnoreAction.command = {
        arguments: [word, match, diagnostic],
        command: "languagetoolLinter.ignoreWordInline",
        title: title,
      };
      wsIgnoreAction.diagnostics = [];
      wsIgnoreAction.diagnostics.push(diagnostic);
      actions.push(wsIgnoreAction);
    }
    return actions;
  }

  private handleAlreadyIgnoredWord(
    word: string,
    diagnostic: LTDiagnostic,
  ): CodeAction[] {
    const actions: CodeAction[] = [];
    if (this.configManager.showIgnoredWordHints()) {
      if (this.configManager.isGloballyIgnoredWord(word)) {
        const actionTitle: string =
          "Remove '" + word + "' from always ignored words.";
        const action: CodeAction = new CodeAction(
          actionTitle,
          CodeActionKind.QuickFix,
        );
        action.command = {
          arguments: [word],
          command: "languagetoolLinter.removeGloballyIgnoredWord",
          title: actionTitle,
        };
        action.diagnostics = [];
        action.diagnostics.push(diagnostic);
        actions.push(action);
      }
      if (this.configManager.isWorkspaceIgnoredWord(word)) {
        const actionTitle: string =
          "Remove '" + word + "' from Workspace ignored words.";
        const action: CodeAction = new CodeAction(
          actionTitle,
          CodeActionKind.QuickFix,
        );
        action.command = {
          arguments: [word],
          command: "languagetoolLinter.removeWorkspaceIgnoredWord",
          title: actionTitle,
        };
        action.diagnostics = [];
        action.diagnostics.push(diagnostic);
      }
    }
    return actions;
  }

  // Get all Rule CodeActions
  private getRuleActions(
    document: TextDocument,
    diagnostic: LTDiagnostic,
  ): CodeAction[] {
    const match: ILanguageToolMatch | undefined = diagnostic.match;
    const actions: CodeAction[] = [];
    if (match) {
      this.getReplacementActions(
        document,
        diagnostic,
        match.replacements,
      ).forEach((action: CodeAction) => {
        actions.push(action);
      });
    }
    return actions;
  }

  // Get all edit CodeActions based on Replacements
  private getReplacementActions(
    document: TextDocument,
    diagnostic: Diagnostic,
    replacements: ILanguageToolReplacement[],
  ): CodeAction[] {
    const actions: CodeAction[] = [];
    replacements.forEach((replacement: ILanguageToolReplacement) => {
      const actionTitle: string = "'" + replacement.value + "'";
      const action: CodeAction = new CodeAction(
        actionTitle,
        CodeActionKind.QuickFix,
      );
      const edit: WorkspaceEdit = new WorkspaceEdit();
      edit.replace(document.uri, diagnostic.range, replacement.value);
      action.edit = edit;
      action.diagnostics = [];
      action.diagnostics.push(diagnostic);
      actions.push(action);
    });
    return actions;
  }

  private readConfigurationFiles(files: string[]) {
    files.forEach((filename: string) => {
      filename = path.resolve(filename); // resolve ~ and relative paths
      if (!fs.existsSync(filename)) {
        Constants.EXTENSION_OUTPUT_CHANNEL.appendLine(
          `Rule-file ${filename} not found.`,
        );
        return;
      }
      this.handleConfigureFile(filename);
    });
  }

  private checkIfRuleIsIgnored(rule: string, token: string, point: Position) {
    const key = rule;
    let key2 = key;
    if (token) {
      key2 = key2 + `(${token})`;
    }
    let result = this.isEnabled(key2, point);
    if (result == undefined) {
      result = this.isEnabled(key, point, false);
    }
    return !result;
  }

  private isEnabled(
    key: string,
    point: Position,
    returnUndefinedOnMiss = true,
  ): boolean | undefined {
    let data: boolean | undefined = this.enabledDisabledRules[key];
    if (data === undefined) {
      data = this.enabledDisabledRulesPerLine[point.line]
        ? this.enabledDisabledRulesPerLine[point.line][key]
        : undefined;
    }
    if (returnUndefinedOnMiss) {
      return data;
    }
    return data === undefined ? true : data;
  }

  private commandMap: Record<string, string> = {
    "ENABLE-FILE": "FILE",
    "DISABLE-FILE": "FILE",
    "DISABLE-LINE": "LINE",
    "ENABLE-LINE": "LINE",
    "DISABLE-NEXT-LINE": "LINE",
    "ENABLE-NEXT-LINE": "LINE",
    "CONFIGURE-FILE": "CONFIG",
  };

  private applyEnableDisable(
    parameter: string,
    enabled: boolean,
    state: Record<string, boolean>,
  ) {
    state = { ...state };
    const trimmed = parameter && parameter.trim();

    let match: RegExpExecArray | null;
    while ((match = this.inlineCommentRuleRe.exec(trimmed))) {
      let key = match[1].toUpperCase();
      if (match[2]) {
        key = key + match[2];
      }
      state[key] = enabled;
    }
    return state;
  }

  private buildRuleList(document: TextDocument): void {
    const lines = document.getText().split("\n");
    let lineIndex = 0;
    if (this.configManager.reloadConfigurationFilesNeeded) {
      Constants.EXTENSION_OUTPUT_CHANNEL.appendLine("reloadConfigurationFiles()");
      this.enabledDisabledRules = {}; // clear current settings
      this.readConfigurationFiles(this.configManager.getConfigurationFiles());
      this.globalEnabledDisabledRules = this.enabledDisabledRules;
      this.configManager.reloadConfigurationFilesNeeded = false;
    }
    this.enabledDisabledRules = this.globalEnabledDisabledRules; // reset
    this.enabledDisabledRulesPerLine = {};
    lines.forEach((line) => {
      let match: RegExpExecArray | null;
      while ((match = this.inlineCommentStartRe.exec(line))) {
        const action = match[2].toUpperCase();
        const startIndex = match.index + match[1].length;
        const endIndex = line.indexOf("-->", startIndex);
        if (endIndex === -1) {
          break;
        }
        const parameter = line.slice(startIndex, endIndex);
        const cmd = this.commandMap[action];
        if (cmd) {
          switch (cmd) {
            case "FILE":
              this.handleEnableDisableFile(action, parameter);
              break;
            case "LINE":
              this.handleEnableDisableLine(action, parameter, lineIndex);
              break;
            case "CONFIG":
              this.handleConfigureFile(parameter);
              break;
          }
        }
      }
      lineIndex++;
    });
    Constants.EXTENSION_OUTPUT_CHANNEL.appendLine(`buildRuleList --> ${Object.keys(this.enabledDisabledRules).length}/${Object.keys(this.enabledDisabledRulesPerLine).length} rules.`);

  }

  private handleEnableDisableFile(action: string, parameter: string) {
    const enabled = action === "ENABLE-FILE";
    this.enabledDisabledRules = this.applyEnableDisable(
      parameter,
      enabled,
      this.enabledDisabledRules,
    );
  }

  private getCwdFromDocument() {
    if (workspace.workspaceFolders !== undefined) {
      const wf = workspace.workspaceFolders[0].uri.path;
      return wf;
    } else {
      const message =
        "YOUR-EXTENSION: Working folder not found, open a folder an try again";

      window.showErrorMessage(message);
      return null;
    }
  }

  private handleConfigureFile(parameter: string) {
    const files = parameter.split(/\s+/);
    const folder = this.getCwdFromDocument() || os.homedir();
    files.forEach((file: string) => {
      if (!file) return;
      const configFile = findUpSync(file, { cwd: folder });
      if (configFile) {
        Constants.EXTENSION_OUTPUT_CHANNEL.appendLine(
          `Reading configuration file ${configFile}`,
        );
        this.loadSingleConfigFile(configFile)
      } else {
        Constants.EXTENSION_OUTPUT_CHANNEL.appendLine(
          `Configuration file ${configFile} not found`,
        );
      }
    });
  }

  private loadSingleConfigFile(configFile: string) {
    const content = fs.readFileSync(configFile, { encoding: "utf8" });
    const lines = content.split("\n");
    const countbefore = Object.keys(this.enabledDisabledRules).length;
    lines.forEach((line: string) => {
      if (/^#/.test(line)) return;
      let match: RegExpExecArray | null;
      while ((match = this.configFileStartRe.exec(line))) {
        if (match.groups) {
          const enabled = match.groups.command.toUpperCase() === "ENABLE";
          this.enabledDisabledRules = this.applyEnableDisable(
            match.groups.parameter,
            enabled,
            this.enabledDisabledRules,
          );
        }
      }
    });
    const countafter = Object.keys(this.enabledDisabledRules).length;
    Constants.EXTENSION_OUTPUT_CHANNEL.appendLine(`Loaded ${countafter - countbefore} items`);
  }

  private handleEnableDisableLine(
    action: string,
    parameter: string,
    lineIndex: number,
  ) {
    const enabled = action === "ENABLE-LINE" || action === "ENABLE-NEXT-LINE";
    const lineno = lineIndex + (action.indexOf("-NEXT-LINE") > 0 ? 1 : 0);
    this.enabledDisabledRulesPerLine[lineno] = this.applyEnableDisable(
      parameter,
      enabled,
      this.enabledDisabledRulesPerLine[lineno] ?? {},
    );
  }
}
