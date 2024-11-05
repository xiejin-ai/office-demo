import React, { useEffect } from 'react';


import loadScript from './load-script';
import IConfig from './config';

declare global {
  interface Window {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    DocsAPI?: any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    DocEditor?: any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    Asc?: { scope: any };
  }
}

interface DocumentEditorProps {
  id: string;

  documentServerUrl: string;

  config: IConfig;

  document_fileType?: string;
  document_title?: string;
  documentType?: string;
  editorConfig_lang?: string;

  type?: string;
  style?: React.CSSProperties;

  onAppReady?: (event: object) => void;
  onDocumentStateChange?: (event: object) => void;
  onMetaChange?: (event: object) => void;
  onDocumentReady?: (event: object) => void;
  onInfo?: (event: object) => void;
  onWarning?: (event: object) => void;
  onError?: (event: object) => void;
  onRequestSharingSettings?: (event: object) => void;
  onRequestRename?: (event: object) => void;
  onMakeActionLink?: (event: object) => void;
  onRequestInsertImage?: (event: object) => void;
  onRequestSaveAs?: (event: object) => void;
  onRequestMailMergeRecipients?: (event: object) => void;
  onRequestCompareFile?: (event: object) => void;
  onRequestEditRights?: (event: object) => void;
  onRequestHistory?: (event: object) => void;
  onRequestHistoryClose?: (event: object) => void;
  onRequestHistoryData?: (event: object) => void;
  onRequestRestore?: (event: object) => void;
  onDownloadAs?: (event: object) => void;
  onRequestCreateNew?: (event: object) => void;
}

export const DocumentEditor = React.memo((props: DocumentEditorProps) => {
  const {
    id,

    documentServerUrl,

    config,

    document_fileType,
    document_title,
    documentType,
    editorConfig_lang,

    type,
    style,

    onAppReady,
    onDocumentStateChange,
    onMetaChange,
    onDownloadAs,
    onDocumentReady,
    onInfo,
    onWarning,
    onError,
    onRequestSharingSettings,
    onRequestRename,
    onMakeActionLink,
    onRequestInsertImage,
    onRequestSaveAs,
    onRequestMailMergeRecipients,
    onRequestCompareFile,
    onRequestEditRights,
    onRequestHistory,
    onRequestHistoryClose,
    onRequestHistoryData,
    onRequestRestore,
    onRequestCreateNew
  } = props;
  // const documentServerUrl = 'http://172.16.200.176:8080/';
  // const documentServerUrl = 'https://only-office-pro.datagrand.com/';

  useEffect(() => {
    if (window?.DocEditor?.instances[id]) {
      window.DocEditor.instances[id].destroyEditor();
      window.DocEditor.instances[id] = undefined;

      console.log('Important props have been changed. Load new Editor.');
      onLoad();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    documentServerUrl,
    // eslint-disable-next-line react-hooks/exhaustive-deps
    JSON.stringify(config),
    document_fileType,
    document_title,
    documentType,
    editorConfig_lang,
    type
  ]);

  useEffect(() => {
    let url = documentServerUrl;
    if (!url.endsWith('/')) url += '/';

    const docApiUrl = `${url}web-apps/apps/api/documents/api.js`;
    loadScript(docApiUrl, 'onlyoffice-api-script')
      .then(() => onLoad())
      .catch((err) => console.error(err));

    return () => {
      if (window?.DocEditor?.instances[id]) {
        window.DocEditor.instances[id].destroyEditor();
        window.DocEditor.instances[id] = undefined;
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const onLoad = () => {
    try {
      if (!window.DocsAPI) throw new Error('DocsAPI is not defined');
      if (window?.DocEditor?.instances[id]) {

        console.log('Skip loading. Instance already exists', id);
        return;
      }

      if (!window?.DocEditor?.instances) {
        window.DocEditor = { instances: {} };
      }

      const initConfig = Object.assign(
        {
          document: {
            fileType: document_fileType,
            title: document_title
          },
          documentType,
          editorConfig: {
            lang: editorConfig_lang
          },
          events: {
            onAppReady: _onAppReady,
            onDocumentStateChange: onDocumentStateChange,
            onMetaChange: onMetaChange,
            onDocumentReady: onDocumentReady,
            onInfo: onInfo,
            onWarning: onWarning,
            onError: onError,
            onRequestSharingSettings: onRequestSharingSettings,
            onRequestRename: onRequestRename,
            onMakeActionLink: onMakeActionLink,
            onRequestInsertImage: onRequestInsertImage,
            onRequestSaveAs: onRequestSaveAs,
            onRequestMailMergeRecipients: onRequestMailMergeRecipients,
            onRequestCompareFile: onRequestCompareFile,
            onRequestEditRights: onRequestEditRights,
            onRequestHistory: onRequestHistory,
            onRequestHistoryClose: onRequestHistoryClose,
            onRequestHistoryData: onRequestHistoryData,
            onRequestRestore: onRequestRestore,
            onDownloadAs: onDownloadAs,
            onRequestCreateNew: onRequestCreateNew
          },
          height: '100%',
          type,
          width: '100%'
        },
        config || {}
      );

      const editor = window.DocsAPI.DocEditor(id, initConfig);
      window.DocEditor.instances[id] = editor;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } catch (err: any) {
      console.error(err);
      onError!(err);
    }
  };

  const _onAppReady = () => {
    onAppReady!(window.DocEditor.instances[id]);
  };

  return (
    <div id={'office-' + id} style={style}>
      <div id={id} style={{ width: '100%', height: '100%' }} />
    </div>
  );
});
