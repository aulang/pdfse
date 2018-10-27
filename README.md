# PDFService
Office格式文件（支持doc、docx、xls、xlsx、ppt、pptx）转换为PDF格式文件、PDF文件签名盖章服务。  
PDF转换使用Office COM组件实现，兼容性好，PDF文件签名盖章基于iText7实现。

### 1. PDF转换
1. 请求地址：/api/pdf/convert
2. 请求参数  
    file: Office格式文件  
    sign：布尔值，true：签名盖章；false：不签名盖章。默认false  
    flag：字符串，盖章标识，支持正则表达式，默认“盖章”
    
### 2. PDF签名
1. 请求地址：/api/pdf/sign
2. 请求参数  
    file: PDF格式文件  
    flag：字符串，盖章标识，支持正则表达式，默认“盖章”
