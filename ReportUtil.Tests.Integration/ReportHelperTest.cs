using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;

namespace ReportUtil.Tests.Integration
{
    [TestClass]
    public class ReportHelperTest
    {
        public class Order
        {
            public DateTime PlaceOrderTime { get; set; }
            public int OrderNumber { get; set; }

            public decimal TotalAmount { get; set; }
            public string RecipientName { get; set; }

            public string Telephone { get; set; }
            public string Address { get; set; }

            public List<OrderDetail> Details { get; set; }
        }

        public class OrderDetail
        {
            public int LineNumber { get; set; }
            public string Name { get; set; }
            public decimal Quantity { get; set; }

            public decimal Price { get; set; }
            public decimal Total { get; set; }

        }

        [TestMethod]
        public void GenerateReportWithTemplate_Test()
        {
            var columnDefs = CreateMasterDetailDataColumns().Where(c => c is ColumnDef<Order>).Cast<ColumnDef<Order>>().ToArray();
            List<Order> orders = GetOrdersData();

            using (MemoryStream targetStream = new MemoryStream())
            {
                var fs = File.OpenRead(Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate", "Test_Order_Template.xlsx"));
                fs.CopyTo(targetStream);
                fs.Close();
                targetStream.Position = 0;

                var stream = new ReportHelper().GenerateReportWithTemplate<Order>(targetStream, orders, columnDefs);

                ValidateSpreadsheetDoc(stream);

                DumpToFile(stream, @"target1.xlsx");
            }
        }

       
        private static List<Order> GetOrdersData()
        {
            return new List<Order>()
            {
                new Order()
                {
                     OrderNumber=1,
                     RecipientName="David",
                     Telephone="1380000000",
                     TotalAmount = 1000000.987654m,
                     Address="FullLink Building， No 23 Chaowai Street, Chaoyang District, Beijing, China ",
                     PlaceOrderTime = DateTime.Now.AddDays(-7)
                },
                new Order()
                {
                    OrderNumber=2,
                    RecipientName="Sherwood",
                    Telephone="1360000000",
                    TotalAmount = 2000000.123456m,
                    Address="Techniq Building， No 5 Zhongguancundong Road, Haidian District, Beijing, China ",
                    PlaceOrderTime = DateTime.Now.AddDays(-1)
                }
            };
        }

        [TestMethod]
        public void GenerateReportWithTemplate_MasterDetail_Test()
        {
            ColumnDefBase[] columns = CreateMasterDetailDataColumns();
            List<Order> orders = CreateOrdersWithDetails();

            using (MemoryStream targetStream = new MemoryStream())
            {
                var fs = File.OpenRead(Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate", "Test_Order_OrderDetail_Template.xlsx"));
                fs.CopyTo(targetStream);
                fs.Close();
                targetStream.Position = 0;

                var stream = new ReportHelper().GenerateReportWithTemplate<Order, OrderDetail>(targetStream, orders, columns, (o) => o.Details);

                ValidateSpreadsheetDoc(stream);

                DumpToFile(stream, @"target2.xlsx");
            }
        }

        [TestMethod]
        public void GenerateReport_MasterDetail_Test()
        {
            ColumnDefBase[] columns = CreateMasterDetailDataColumns();
            List<Order> orders = CreateOrdersWithDetails();

            var stream = new ReportHelper().GenerateReport<Order, OrderDetail>(orders, columns, (o) => o.Details);
            ValidateSpreadsheetDoc(stream);
            DumpToFile(stream, @"target3.xlsx");

        }

        [TestMethod]
        public void GenerateReport_Test()
        {
            var columnDefs = CreateMasterDetailDataColumns().Where(c => c is ColumnDef<Order>).Cast<ColumnDef<Order>>().ToArray();
            List<Order> orders = GetOrdersData();

            var stream = new ReportHelper().GenerateReport<Order>(orders, columnDefs);
            ValidateSpreadsheetDoc(stream);
            DumpToFile(stream, @"target4.xlsx");

        }

        [TestMethod]
        public void GenerateMultipSheetReport_Test()
        {
            var columnDefs = CreateMasterDetailDataColumns().Where(c => c is ColumnDef<Order>).Cast<ColumnDef<Order>>().ToArray();
            List<Order> orders = GetOrdersData();
            var reportHelper = new ReportHelper();
            var stream = reportHelper.GenerateReport(orders, columnDefs);
           
            reportHelper.GenerateReportWithTemplate(stream, orders, columnDefs, "Sheet2");
            ValidateSpreadsheetDoc(stream);

            DumpToFile(stream, @"target5.xlsx");

        }

        private static void ValidateSpreadsheetDoc(Stream stream)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, false))
            {
                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(doc);

                foreach (var error in errors)
                {
                    Console.WriteLine("Error description: {0}", error.Description);
                    Console.WriteLine("Content type of part with error: {0}",
                        error.Part.ContentType);
                    Console.WriteLine("Location of error: {0}", error.Path.XPath);
                }

                Assert.AreEqual(errors.Count(), 0);
            }
        }

        private static List<Order> CreateOrdersWithDetails()
        {
            return new List<Order>()
            {
                new Order()
                {
                     OrderNumber=1,
                     RecipientName="David",
                     Telephone="1380000000",
                     TotalAmount = 1000000.987654m,
                     Address="FullLink Building， No 23 Chaowai Street, Chaoyang District, Beijing, China ",
                     PlaceOrderTime = DateTime.Now.AddDays(-7),
                     Details = new List<OrderDetail>()
                     {
                         new OrderDetail()
                         {
                            LineNumber=1,
                            Name="CPU",
                            Price=2000.1234m,
                            Quantity = 2,
                            Total=4000.2468m

                         },
                          new OrderDetail()
                         {
                            LineNumber=2,
                            Name="Monitor",
                            Price=1000.1234m,
                            Quantity = 1,
                            Total=1000.1234m
                         },
                           new OrderDetail()
                         {
                            LineNumber=3,
                            Name="HardDrive",
                            Price=2000.1234m,
                            Quantity = 1,
                            Total=2000.1234m
                         }
                     }
                },
                new Order()
                {
                    OrderNumber=2,
                    RecipientName="Sherwood",
                    Telephone="1360000000",
                    TotalAmount = 2000000.123456m,
                    Address="Techniq Building， No 5 Zhongguancundong Road, Haidian District, Beijing, China ",
                    PlaceOrderTime = DateTime.Now.AddDays(-1),
                     Details = new List<OrderDetail>()
                     {
                         new OrderDetail()
                         {
                            LineNumber=1,
                            Name="处理器",
                            Price=2000.1234m,
                            Quantity = 2,
                            Total=4000.2468m

                         },
                          new OrderDetail()
                         {
                            LineNumber=2,
                            Name="显示器",
                            Price=1000.1234m,
                            Quantity = 1,
                            Total=1000.1234m
                         },
                           new OrderDetail()
                         {
                            LineNumber=3,
                            Name="硬盘",
                            Price=2000.1234m,
                            Quantity = 1,
                            Total=2000.1234m
                         }
                     }
                }
            };
        }

        private static ColumnDefBase[] CreateMasterDetailDataColumns()
        {
            return new ColumnDefBase[]
                        {
                new ColumnDef<Order>()
                {
                    Captain = "下单时间",
                    GetValueFunc = (o) => new CellValue(o.PlaceOrderTime.ToString("s")),
                    TargetDataType = new EnumValue<CellValues>(CellValues.Date)
                },
                 new ColumnDef<Order>()
                {
                    Captain = "订单号",
                    GetValueFunc = (o) => new CellValue(o.OrderNumber.ToString()),
                    TargetDataType = new EnumValue<CellValues>(CellValues.String)
                },

                new ColumnDef<OrderDetail>()
                {
                    Captain = "序号",
                    GetValueFunc=(o)=>new CellValue(o.LineNumber.ToString()),
                    TargetDataType=new EnumValue<CellValues>(CellValues.Number)
                },
                new ColumnDef<OrderDetail>()
                {
                    Captain = "名称",
                    GetValueFunc=(o)=>new CellValue(o.Name),
                    TargetDataType=new EnumValue<CellValues>(CellValues.String)
                },
                new ColumnDef<OrderDetail>()
                {
                    Captain = "数量",
                    GetValueFunc=(o)=>new CellValue(o.Quantity.ToString()),
                    TargetDataType=new EnumValue<CellValues>(CellValues.Number)
                },
                new ColumnDef<OrderDetail>()
                {
                    Captain = "单价",
                    GetValueFunc=(o)=>new CellValue(o.Price.ToString()),
                    TargetDataType=new EnumValue<CellValues>(CellValues.Number)
                },
                new ColumnDef<OrderDetail>()
                {
                    Captain = "小计",
                    GetValueFunc=(o)=>new CellValue(o.Total.ToString()),
                    TargetDataType=new EnumValue<CellValues>(CellValues.Number)
                },

                new ColumnDef<Order>()
                {
                    Captain = "收货人",
                    GetValueFunc = (o) => new CellValue(o.RecipientName),
                    TargetDataType = new EnumValue<CellValues>(CellValues.String)
                },
                 new ColumnDef<Order>()
                {
                    Captain = "电话",
                    GetValueFunc = (o) => new CellValue(o.Telephone),
                    TargetDataType = new EnumValue<CellValues>(CellValues.String)
                },
                new ColumnDef<Order>()
                {
                    Captain = "送货地址",
                    GetValueFunc = (o) => new CellValue(o.Address),
                    TargetDataType = new EnumValue<CellValues>(CellValues.String)
                },
                 new ColumnDef<Order>()
                {
                    Captain = "总价",
                    GetValueFunc = (o) => new CellValue(o.TotalAmount.ToString()),
                    TargetDataType = new EnumValue<CellValues>(CellValues.Number)
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
