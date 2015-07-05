    /*
    *      c-xml.c
    */

    #include <windows.h>
    #include <ole2.h>

    #pragma comment(linker, "/entry:Start")
    #pragma comment(linker, "/subsystem:windows")
    #pragma comment(lib, "kernel32.lib")
    #pragma comment(lib, "user32.lib")
    #pragma comment(lib, "ole32.lib")
    #pragma comment(lib, "oleaut32.lib")

    void Start(){
       IXMLDOMDocument*   pXMLDoc;
       IXMLDOMNode*      pXMLNode;
       IXMLDOMNode*      pXMLChildNode;
       IXMLDOMElement*      pXMLElement;
       VARIANT_BOOL bSucc;
       VARIANT var;
       BSTR bs;

       CoInitialize(0);

       CoCreateInstance(&CLSID_DOMDocument, NULL, CLSCTX_INPROC_SERVER,
          &IID_IXMLDOMDocument, (void**) &pXMLDoc);

       pXMLDoc->lpVtbl->loadXML(pXMLDoc, L"<Node><SubNode/></Node>", &bSucc);

       pXMLDoc->lpVtbl->selectSingleNode(pXMLDoc, L"Node/SubNode", &pXMLNode);

       var.vt = VT_INT;   var.intVal = NODE_ELEMENT;
       pXMLDoc->lpVtbl->createNode(pXMLDoc, var, L"ChildNode", NULL, &pXMLChildNode);

       pXMLNode->lpVtbl->appendChild(pXMLNode, pXMLChildNode, &pXMLNode);

       pXMLNode->lpVtbl->QueryInterface(pXMLNode, &IID_IXMLDOMElement, (void**) &pXMLElement);

       var.vt = VT_BSTR;   var.bstrVal = SysAllocString(L"c");
       pXMLElement->lpVtbl->setAttribute(pXMLElement, L"Language", var);
       SysFreeString(var.bstrVal);

       var.vt = VT_BSTR;   var.bstrVal = SysAllocString(L"c.xml");
       pXMLDoc->lpVtbl->save(pXMLDoc, var);
       SysFreeString(var.bstrVal);

       pXMLDoc->lpVtbl->get_xml(pXMLDoc, &bs);
       MessageBoxW(0, bs, L"get_xml", 0);

       CoUninitialize();

       ExitProcess(0);
    }
