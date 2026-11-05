export interface MemoItem {
  name: string;
  path: string;
  /** 下载文件名；默认 `${name}.md` */
  downloadName?: string;
}

export interface MemoSearchSource {
  name: string;
  path: string;
  markdown: string;
}

export interface MemoSearchResult extends MemoSearchSource {
  snippet: string;
}
