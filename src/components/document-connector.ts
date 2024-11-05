export class DocumentConnector {
  frame?: HTMLIFrameElement;
  guid: string;
  isConnected: boolean;
  callbacks: Function[];
  events: any;
  tasks: any[];

  constructor() {
    this.guid = 'asc.{' + this.generateGuid() + '}';
    this.isConnected = false;
    this.callbacks = [];
    this.events = {};
    this.tasks = [];
  }

  generateGuid(): string {
    const a = (): string =>
      Math.floor(65536 * (1 + Math.random()))
        .toString(16)
        .substring(1);
    return a() + a() + '-' + a() + '-' + a() + '-' + a() + '-' + a() + a() + a();
  }

  onMessage = (event: any) => {
    if (typeof event.data === 'string') {
      let obj: any = {};
      try {
        obj = JSON.parse(event.data);
      } catch (c) {
        obj = {};
      }
      if (obj?.type !== 'onExternalPluginMessageCallback') {
        return;
      }
      const data = obj.data;

      if (data.guid === this.guid) {
        switch (data.type) {
          case 'onMethodReturn':
            if (this.callbacks.length > 0) {
              const callback = this.callbacks.shift();
              if (callback) {
                callback(data.methodReturnData);
              }
            }
            if (this.tasks.length > 0) {
              this.sendMessage(this.tasks.shift());
            }
            break;
          case 'onCommandCallback':
            if (this.callbacks.length > 0) {
              const callback = this.callbacks.shift();
              if (callback) {
                callback(data.commandReturnData);
              }
            }
            if (this.tasks.length > 0) {
              this.sendMessage(this.tasks.shift());
            }
            break;
          case 'onEvent':
            if (data.eventName && this.events[data.eventName]) {
              this.events[data.eventName](data.eventData);
            }
            break;
        }
      }
    }
  };

  sendMessage(data: any): void {
    const message = {
      frameEditorId: 'iframeEditor',
      type: 'onExternalPluginMessage',
      subType: 'connector',
      data: { ...data, guid: this.guid }
    };

    this.frame?.contentWindow?.postMessage(JSON.stringify(message), '*');
  }

  connect(id: string): void {
    this.frame = document.querySelector(`#office-${id} iframe`) as HTMLIFrameElement;
    if (this.isConnected) {
      console.log('This connector is already connected');
    } else {
      if (window.addEventListener) {
        window.addEventListener('message', this.onMessage, false);
      }
      this.isConnected = true;
      this.sendMessage({ type: 'register' });
    }
  }

  disconnect(): void {
    if (this.isConnected) {
      if (window.removeEventListener) {
        window.removeEventListener('message', this.onMessage, false);
      }
      this.isConnected = false;
      this.sendMessage({ type: 'unregister' });
    } else {
      console.log('This connector is already disconnected');
    }
  }
  /**
   * 执行的命令没有上下文 默认会有一个 Api 的 内部变量
   * const params = {};
   * window.Asc!.scope = params;
   * this.connector.callCommand<any>(
   *   () => {
   *     const WordApi = Api as WordApi;
   *     const scope = Asc!.scope as typeof params;
   *   },
   *   (data) => {
   *    console.log(data);
   *   }
   * );
   */
  callCommand<T = any>(fun: () => T, params?: Object, callback: (data: T) => void = () => {}, recalculate: boolean = true): Promise<T> {
    return new Promise((resolve) => {
      const func = (data: T) => {
        callback(data);
        resolve(data);
      };
      if (this.isConnected) {
        this.callbacks.push(func);
        const aStr = 'var Asc = {}; Asc.scope = ' + JSON.stringify(params || {}) + '; var scope = Asc.scope; (' + fun.toString() + ')();';
        const cObj = { type: 'command', recalculate: recalculate, data: aStr };
        if (this.callbacks.length !== 1) {
          this.tasks.push(cObj);
        } else {
          this.sendMessage(cObj);
        }
      } else {
        console.log('Connector is not connected with editor');
      }
    });
  }

  executeMethod(methodName: string, params: any, fun?: Function): void {
    if (this.isConnected) {
      this.callbacks.push(fun ?? function () {});
      const aObj = { type: 'method', methodName: methodName, data: params };
      if (this.callbacks.length !== 1) {
        this.tasks.push(aObj);
      } else {
        this.sendMessage(aObj);
      }
    } else {
      console.log('Connector is not connected with editor');
    }
  }

  attachEvent(methodName: string, fun: Function): void {
    if (this.isConnected) {
      this.events[methodName] = fun;
      this.sendMessage({ type: 'attachEvent', name: methodName });
    } else {
      console.log('Connector is not connected with editor');
    }
  }

  detachEvent(methodName: string): void {
    if (this.events[methodName]) {
      delete this.events[methodName];
      if (this.isConnected) {
        this.sendMessage({ type: 'detachEvent', name: methodName });
      } else {
        console.log('Connector is not connected with editor');
      }
    }
  }
}
// 这是一个 onlyOffice 的内部对象，只是为了补充类型防止报错
// TODO: 之后可以考虑把这个对象的类型补充完整
export interface WordApi {
  GetDocument: () => ApiDocument;
  RemoveSelection: () => void;
  AddComment: Function;
  attachEvent: Function;
  ConvertDocument: Function;
  CreateBlipFill: Function;
  CreateBlockLvlSdt: Function;
  CreateBullet: Function;
  CreateChart: Function;
  CreateGradientStop: Function;
  CreateHyperlink: Function;
  CreateImage: Function;
  CreateInlineLvlSdt: Function;
  CreateLinearGradientFill: Function;
  CreateNoFill: Function;
  CreateNumbering: Function;
  CreateOleObject: Function;
  CreateParagraph: Function;
  CreatePatternFill: Function;
  CreatePresetColor: Function;
  CreateRadialGradientFill: Function;
  CreateRange: Function;
  CreateRGBColor: Function;
  CreateRun: Function;
  CreateSchemeColor: Function;
  CreateShape: Function;
  CreateSolidFill: Function;
  CreateStroke: Function;
  CreateTable: Function;
  CreateTextPr: Function;
  CreateWordArt: Function;
  detachEvent: Function;
  FromJSON: Function;
  GetFullName: Function;
  GetMailMergeReceptionsCount: Function;
  GetMailMergeTemplateDocContent: Function;
  LoadMailMergeData: Function;
  MailMerge: Function;
  ReplaceDocumentContent: Function;
  ReplaceTextSmart: Function;
  Save: Function;
  [key: string]: Function;
}

export interface ApiDocument {
  AcceptAllRevisionChanges: Function;
  AddComment: Function;
  AddElement: Function;
  AddEndnote: Function;
  AddFootnote: Function;
  AddTableOfContents: Function;
  AddTableOfFigures: Function;
  ClearAllFields: Function;
  CreateNewHistoryPoint: Function;
  CreateNumbering: Function;
  CreateSection: Function;
  CreateStyle: Function;
  DeleteBookmark: Function;
  GetAllBookmarksNames: () => string[];
  GetAllCaptionParagraphs: Function;
  GetAllCharts: Function;
  GetAllComments: Function;
  GetAllContentControls: Function;
  GetAllDrawingObjects: Function;
  GetAllForms: Function;
  GetAllHeadingParagraphs: Function;
  GetAllImages: Function;
  GetAllNumberedParagraphs: Function;
  GetAllOleObjects: Function;
  GetAllParagraphs: Function;
  GetAllShapes: Function;
  GetAllStyles: Function;
  GetAllTables: Function;
  GetAllTablesOnPage: Function;
  GetBookmarkRange: (name: string) => ApiRange;
  GetClassType: Function;
  GetCommentById: Function;
  GetCommentsReport: Function;
  GetContent: Function;
  GetContentControlsByTag: Function;
  GetDefaultParaPr: Function;
  GetDefaultStyle: Function;
  GetDefaultTextPr: Function;
  GetElement: Function;
  GetElementsCount: Function;
  GetEndNotesFirstParagraphs: Function;
  GetFinalSection: Function;
  GetFootnotesFirstParagraphs: Function;
  GetFormsByTag: Function;
  GetPageCount: Function;
  GetRange: (start: number, end: number) => ApiRange;
  GetRangeBySelect: Function;
  GetReviewReport: Function;
  GetSections: Function;
  GetSelectedDrawings: Function;
  GetStatistics: Function;
  GetStyle: Function;
  GetTagsOfAllContentControls: Function;
  GetTagsOfAllForms: Function;
  GetWatermarkSettings: Function;
  InsertContent: Function;
  InsertWatermark: Function;
  IsTrackRevisions: Function;
  Last: Function;
  Push: Function;
  RejectAllRevisionChanges: Function;
  RemoveAllElements: Function;
  RemoveElement: Function;
  RemoveSelection: Function;
  RemoveWatermark: Function;
  ReplaceCurrentImage: Function;
  ReplaceDrawing: Function;
  Search: Function;
  SearchAndReplace: Function;
  SetControlsHighlight: Function;
  SetEvenAnd: Function;
  SetFormsHighlight: Function;
  SetTrackRevisions: Function;
  SetWatermarkSettings: Function;
  ToHtml: Function;
  ToJSON: Function;
  ToMarkdown: Function;
  UpdateAllTOC: Function;
  UpdateAllTOF: Function;
  GetCursorPosition: () => number;
  MoveCursor: (direction: 'Left' | 'Right' | 'Up' | 'Down') => void;
  [key: string]: Function;
}

export interface ApiRange {
  AddBookmark: Function;
  AddComment: Function;
  AddHyperlink: Function;
  AddText: Function;
  Delete: Function;
  ExpandTo: (param: ApiRange) => ApiRange;
  GetAllParagraphs: () => ApiParagraph[];
  GetClassType: Function;
  GetParagraph: Function;
  GetStartAndEnd: () => [number, number];
  GetRange: Function;
  GetText: Function;
  IntersectWith: Function;
  Select: Function;
  SetBold: Function;
  SetCaps: Function;
  SetColor: Function;
  SetDoubleStrikeout: Function;
  SetFontFamily: Function;
  SetFontSize: Function;
  SetHighlight: (params: string | { r: number; g: number; b: number }) => void;
  SetItalic: Function;
  SetPosition: Function;
  SetShd: Function;
  SetSmallCaps: Function;
  SetSpacing: Function;
  SetStrikeout: Function;
  SetStyle: Function;
  SetTextPr: Function;
  SetUnderline: Function;
  SetVertAlign: Function;
  ToJSON: Function;

  [key: string]: Function;
}

export interface ApiParagraph {
  AddBookmarkCrossRef: Function;
  AddCaption: Function;
  AddCaptionCrossRef: Function;
  AddColumnBreak: Function;
  AddComment: Function;
  AddDrawing: Function;
  AddElement: Function;
  AddEndnoteCrossRef: Function;
  AddFootnoteCrossRef: Function;
  AddHeadingCrossRef: Function;
  AddHyperlink: Function;
  AddInlineLvlSdt: Function;
  AddLineBreak: Function;
  AddNumberedCrossRef: Function;
  AddPageBreak: Function;
  AddPageNumber: Function;
  AddPagesCount: Function;
  AddTabStop: Function;
  AddText: Function;
  Copy: Function;
  Delete: Function;
  GetAllCharts: Function;
  GetAllContentControls: Function;
  GetAllDrawingObjects: Function;
  GetAllImages: Function;
  GetAllOleObjects: Function;
  GetAllShapes: Function;
  GetClassType: Function;
  GetElement: Function;
  GetElementsCount: Function;
  GetFontNames: Function;
  GetIndFirstLine: Function;
  GetIndLeft: Function;
  GetIndRight: Function;
  GetJc: Function;
  GetLastRunWithText: Function;
  GetNext: Function;
  GetNumbering: Function;
  GetParagraphMarkTextPr: Function;
  GetParaPr: Function;
  GetParentContentControl: Function;
  GetParentTable: Function;
  GetParentTableCell: Function;
  GetPosInParent: Function;
  GetPrevious: Function;
  GetRange: (start?: number, end?: number) => ApiRange;
  GetSection: Function;
  GetShd: Function;
  GetSpacingAfter: Function;
  GetSpacingBefore: Function;
  GetSpacingLineRule: Function;
  GetSpacingLineValue: Function;
  GetStyle: Function;
  GetText: Function;
  GetTextPr: Function;
  InsertInContentControl: Function;
  InsertParagraph: Function;
  Last: Function;
  Push: Function;
  RemoveAllElements: Function;
  RemoveElement: Function;
  ReplaceByElement: Function;
  Search: Function;
  Select: Function;
  SetBetweenBorder: Function;
  SetBold: Function;
  SetBottomBorder: Function;
  SetBullet: Function;
  SetCaps: Function;
  SetColor: Function;
  SetContextualSpacing: Function;
  SetDoubleStrikeout: Function;
  SetFontFamily: Function;
  SetFontSize: Function;
  SetHighlight: (params: string | { r: number; g: number; b: number }) => void;
  SetIndFirstLine: Function;
  SetIndLeft: Function;
  SetIndRight: Function;
  SetItalic: Function;
  SetJc: Function;
  SetKeepLines: Function;
  SetKeepNext: Function;
  SetLeftBorder: Function;
  SetNumbering: Function;
  SetNumPr: Function;
  SetPageBreakBefore: Function;
  SetPosition: Function;
  SetRightBorder: Function;
  SetSection: Function;
  SetShd: Function;
  SetSmallCaps: Function;
  SetSpacing: Function;
  SetSpacingAfter: Function;
  SetSpacingBefore: Function;
  SetSpacingLine: Function;
  SetStrikeout: Function;
  SetStyle: Function;
  SetTabs: Function;
  SetTextPr: Function;
  SetTopBorder: Function;
  SetUnderline: Function;
  SetVertAlign: Function;
  SetWidowControl: Function;
  ToJSON: Function;
  WrapInMailMergeField: Function;
  [key: string]: Function;
}
