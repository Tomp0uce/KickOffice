declare module 'markdown-it-task-lists';
declare module 'markdown-it-deflist';
declare module 'markdown-it-footnote';
declare module 'turndown' {
  interface Options {
    headingStyle?: 'setext' | 'atx';
    bulletListMarker?: '-' | '+' | '*';
    codeBlockStyle?: 'indented' | 'fenced';
    fence?: string;
  }
  interface Rule {
    filter: string | string[] | ((node: any) => boolean);
    replacement: (content: string, node: any, options: Options) => string;
  }
  class TurndownService {
    constructor(options?: Options);
    turndown(html: string): string;
    addRule(key: string, rule: Rule): this;
    use(plugin: (service: TurndownService) => void): this;
  }
  export = TurndownService;
}
declare module 'diff-match-patch' {
  class diff_match_patch {
    diff_main(text1: string, text2: string): [number, string][];
    diff_cleanupSemantic(diffs: [number, string][]): void;
  }
  export = diff_match_patch;
}
