# ReportUtil

[![NuGet](https://img.shields.io/nuget/v/ReportUtil.svg)](https://www.nuget.org/packages/ReportUtil)
[![NuGet](https://img.shields.io/nuget/dt/ReportUtil.svg)](https://www.nuget.org/packages/ReportUtil)
[![Build Status](https://dev.azure.com/zhangshangwu0605/zhangshangwu/_apis/build/status/zhangshangwu.ReportUtil?branchName=master)](https://dev.azure.com/zhangshangwu0605/zhangshangwu/_build/latest?definitionId=1&branchName=master)

A  .net core utility library to export generic lists to excel file using [Open Xml SDK](https://github.com/OfficeDev/Open-XML-SDK). The collections could be master-detail style, and the cells for master fields could be merged. 

Nuget package is available [here](https://www.nuget.org/packages/ReportUtil/).

The result snapshot looks like as following.



![](https://github.com/zhangshangwu/ReportUtil/blob/master/Snapshot.PNG?s=300)

## How to use

1. Install the library from Nuget

```
> Install-Package ReportUtil
```
2. Suppose we have following model

```
public class Category
    {
        public string Code { get; set; }
        public string Name { get; set; }

        public List<Product> Products { get; set; }
    }


public class Product
    {
        public int ProductId { get; set; }
        public string Name { get; set; }

        public string Desctiption { get; set; }

        public decimal Price { get; set; }

        public DateTime CreateDate { get; set; }

    }
    
```

and We have a collection of Category.

2.1 Firstly, define the ColumnDef array,

```
var columnDefs = new ColumnDefBase[]
                        {
                            new ColumnDef<Category>()
                            {
                                 Captain="CategoryCode",
                                 TargetDataType=new  EnumValue<CellValues>(CellValues.String),
                                 GetValueFunc=(o)=>new  CellValue(o.Code)
                            },
                            new ColumnDef<Category>()
                            {
                                 Captain="CategoryName",
                                 TargetDataType=new  EnumValue<CellValues>(CellValues.String),
                                 GetValueFunc=(o)=>new  CellValue(o.Name)
                            },
                            new ColumnDef<Product>()
                            {
                                 Captain="ProductId",
                                 TargetDataType=new  EnumValue<CellValues>(CellValues.String),
                                 GetValueFunc=(o)=>new  CellValue(o.ProductId.ToString())
                            },
                            
                            ...//ommit some code for simplicity
                            new ColumnDef<Product>()
                            {
                                Captain="CreateDate",
                                TargetDataType=new  EnumValue<CellValues>(CellValues.Date),
                                GetValueFunc=(o)=>new  CellValue(o.CreateDate.ToString("s"))
                            }
                        };
                            
```

2.2 If we have already a template xlsx file to define the captain rows and read it into stream, call the function like following,

```
new ReportHelper().GenerateReportWithTemplate<Category, Product>(stream,data, columnDefs, (o) => o.Products);
```
2.3 If there's no template file, just call another function ,

```
var stream = new ReportHelper().GenerateReport<Category, Product>(data, columnDefs, (o) => o.Products);
```

2.4 Save the stream to file or a http response output stream, you will get an excel report file.

Enjoy it!
