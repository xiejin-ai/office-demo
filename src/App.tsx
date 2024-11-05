import { useMemo, useRef, useState } from "react";
import "./App.css";
import { DocumentConnector, DocumentEditor, WordApi } from "./components";

import IConfig from "./components/document-editor/config";

// 这是访问onlyoffice服务的地址 如果映射到本地8081端口则为 http://localhost:8081/
const DocumentServerUrl = "http://218.78.212.10:30080/";
// const DocumentServerUrl = "http://192.168.1.163:4040/"
interface IComment {
  Id: string;
  Data: { Text: string; UserName: string };
}
const id = new Date().getTime().toString()

function App() {
  const connectorRef = useRef(new DocumentConnector());
  const [bookmarks, setBookmarks] = useState<string[]>([]);

  const onDocumentReady = async () => {
    console.log(1111) 
    connectorRef.current.connect(id);
    const connector = connectorRef.current!;
    const arr = await connector.callCommand(() => {
      // debugger;
      // @ts-ignore
      const wordApi = Api as WordApi;
     const doc = wordApi.GetDocument()
     const ps= doc.GetAllParagraphs().map(p=>p.GetText())
      return ps 
    });
    console.log(arr,1111)
    
  };

  const onInject = () => {
    const connector = connectorRef.current!;
    const params = {
      bookmarks
    }
    connector.callCommand(() => {
      // @ts-ignore
      const wordApi = Api as WordApi;
     //@ts-ignore
     const { bookmarks } = Asc!.scope as typeof params;
     const oDoc = wordApi.GetDocument();
     const range = oDoc.GetBookmarkRange(bookmarks[bookmarks.length-1]);
     debugger;
     range.Select();
      
    },params);
  };

  const onJump = () => {
    const connector = connectorRef.current!;
    const params = {
      bookmarks
    }
    connector.callCommand(() => {
      // @ts-ignore
      const wordApi = Api as WordApi;
     //@ts-ignore
     const { bookmarks } = Asc!.scope as typeof params;
     const oDoc = wordApi.GetDocument();
     const range = oDoc.GetRange(300000,300004);
    //  range.SetHighlight('darkGrey');
    
    
     range.Select();
      
    },params);
  };
  const onDownloadAs = (e:any)=>{
    console.log(1111,e)
  }

  const goTo = ()=>{
    const connector = connectorRef.current!;

    connector.executeMethod('GetAllContentControls',null,(data:any[])=>{
    
      // debugger;
      connector.executeMethod("SelectContentControl",[data[data.length-2].InternalId])
    })
    
  }
const [value,setValue] = useState('')
const [url,setUrl] = useState("https://dev-microcraft-storage.microware1985.cn/microcraft/upload_files/b3cb25f3-089a-46a8-8db5-3bb9c4a89aa3/790c2fd4-3164-485a-ac16-0ed73ee2f019.pdf?AWSAccessKeyId=zF0oScWIYKm7Wj82u1Y8&Signature=IrVKTjsUw%2BindFq60Uvb%2F8ivyEA%3D&Expires=1730775995"
)
const [key,setKey] = useState('')
const [fileType,setFileType] = useState('pdf')
const [documentType,setDocumentType] = useState('word')

  const config: IConfig = useMemo(()=>{
    return {
      document: {
        fileType: fileType,
        key: key, // 更换文件需要更换key
        title: "Test." + fileType,
        url: url, // 这是word文件的地址 需要换成自己本地的ip 需要保证onlyoffice服务能访问到这个地址
        permissions: {
          edit: true,
        
        },
       
      },
      documentType: documentType,
      editorConfig: {
        lang: "zh-ch",
        callbackUrl: "http://www.baidu.com",
        user: {
          id: "11",
          name: "111",
        },
      },
    };
  
  },[url,key])
  return (
    <div className="App">
      <DocumentEditor
        style={{ flex: 1 }}
        id={id}
        documentServerUrl={DocumentServerUrl}
        config={config}
        onDocumentReady={onDocumentReady}
        onDownloadAs={onDownloadAs}
        onRequestCreateNew={()=>{
          console.log(1111)
        }}
        
     
      />
      <div style={{ width: "200px",display:'flex',flexDirection:'column',gap:'10px' }}>
       <input type="text"  placeholder="文件地址" onChange={(e)=>{
       setValue(e.target.value)  
       }}/>
     
       <input type="text"  placeholder="文件类型" onChange={(e)=>{
        setFileType(e.target.value)
       }}/>
       <input type="text"  placeholder="文档类型" onChange={(e)=>{
        setDocumentType(e.target.value)
      }}/>
      <button onClick={()=>{
       setUrl(value)
       setKey(new Date().getTime().toString())
      }}>访问</button>

      </div>
    </div>
  );
}

export default App;

// connector.callCommand(function () {
//   var oDocument = Api.GetDocument();
//   console.log(oDocument);
//   oDocument.SelectComment();
//   var oParagraph = Api.CreateParagraph();
//   oParagraph.AddText("Hello world!");
//   oDocument.InsertContent([oParagraph]);
// }, true);
