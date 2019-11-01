using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ReportUtil;
using System;
using System.Collections.Generic;
using System.IO;

namespace RepoUtilSample
{
    class Program
    {
        static void Main(string[] args)
        {

            List<Category> data = Repository.GetCategories();

            ColumnDefBase[] columnDefs = CreateMasterDetailDataColumns();

           
            var reportHelper = new ReportHelper();

            MemoryStream targetStream = new MemoryStream();
            var fs = File.OpenRead(Path.Combine(Directory.GetCurrentDirectory(), "Template", "ProductsWithCategories.xlsx"));

            fs.CopyTo(targetStream);
            fs.Close();
            targetStream.Position = 0;
            reportHelper.GenerateReportWithTemplate<Category, Product>(targetStream,data,columnDefs,(o)=>o.Products);

            //If there's no template, use following code
            //var targetStream = reportHelper.GenerateReport<Category, Product>(data, columnDefs, (o) => o.Products);

            DumpToFile(targetStream, "target1.xlsx");
            Console.WriteLine("Reporte generated successfully!");
            Console.ReadKey();
        }

        private static ColumnDefBase[] CreateMasterDetailDataColumns()
        {
            return new ColumnDefBase[]
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
                            new ColumnDef<Product>()
                            {
                                 Captain="ProductName",
                                 TargetDataType=new  EnumValue<CellValues>(CellValues.String),
                                 GetValueFunc=(o)=>new  CellValue(o.Name)
                            },
                            new ColumnDef<Product>()
                            {
                                 Captain="Description",
                                 TargetDataType=new  EnumValue<CellValues>(CellValues.String),
                                 GetValueFunc=(o)=>new  CellValue(o.Desctiption)
                            },
                            new ColumnDef<Product>()
                            {
                                 Captain="Price",
                                 TargetDataType=new  EnumValue<CellValues>(CellValues.Number),
                                 GetValueFunc=(o)=>new  CellValue(o.Price.ToString())
                            },
                            new ColumnDef<Product>()
                            {
                                Captain="CreateDate",
                                TargetDataType=new  EnumValue<CellValues>(CellValues.Date),
                                GetValueFunc=(o)=>new  CellValue(o.CreateDate.ToString("s"))
                            }
                        };
        }

        private static void DumpToFile(Stream stream, string fileName)
        {
            stream.Position = 0;
            if (File.Exists(fileName)) File.Delete(fileName);
            using (FileStream targetfile = new FileStream(fileName, FileMode.Create))
            {
                stream.CopyTo(targetfile);
            }

        }
    }
}
