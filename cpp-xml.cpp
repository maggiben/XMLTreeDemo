    //
    //      cpp-xml.cpp
    //

    #include <windows.h>
    #include <ole2.h>
    #include <comutil.h>

    #pragma comment(linker, "/entry:Start")
    #pragma comment(linker, "/subsystem:windows")
    #pragma comment(lib, "kernel32.lib")
    #pragma comment(lib, "user32.lib")
    #pragma comment(lib, "ole32.lib")
    #pragma comment(lib, "comsupp.lib")

    void Start(){
       IXMLDOMDocument*   pXMLDoc;
       IXMLDOMNode*      pXMLNode;
       IXMLDOMNode*      pXMLChildNode;
       IXMLDOMElement*      pXMLElement;
       VARIANT_BOOL bSucc;
       BSTR bs;

       CoInitialize(0);

       CoCreateInstance( CLSID_DOMDocument, NULL, CLSCTX_INPROC_SERVER,
           IID_IXMLDOMDocument, (void**) &pXMLDoc);

       pXMLDoc->loadXML(L"<Node><SubNode/></Node>", &bSucc);

       pXMLDoc->selectSingleNode(L"Node/SubNode", &pXMLNode);

       pXMLDoc->createNode(_variant_t(NODE_ELEMENT), L"ChildNode", NULL, &pXMLChildNode);

       pXMLNode->appendChild(pXMLChildNode, &pXMLNode);

       pXMLNode->QueryInterface(IID_IXMLDOMElement, (void**) &pXMLElement);

       pXMLElement->setAttribute(L"Language", _variant_t(L"c++"));

       pXMLDoc->save(_variant_t(L"c++.xml"));

       pXMLDoc->get_xml(&bs);

       MessageBoxW(0, bs, L"get_xml", 0);

       CoUninitialize();

       ExitProcess(0);
    }
