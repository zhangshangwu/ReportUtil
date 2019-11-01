using System;
using System.Collections.Generic;
using System.Text;

namespace RepoUtilSample
{
   public class Repository
    {

        public static List<Category> GetCategories()
        {
            return new List<Category>()
            {
                new Category()
                {
                    Code="Category1",
                    Name="Food",
                    Products=new List<Product>()
                    {
                        new Product()
                        {
                            Name="Apple",
                            Desctiption="An apple is a sweet, edible fruit produced by an apple tree (Malus domestica) ",
                            Price=5.00m,
                            ProductId=1,
                            CreateDate=DateTime.Now.AddMonths(-5)
                        },
                        new Product()
                        {
                            Name="Pizza",
                            Desctiption="a savory dish of Italian origin, consisting of a usually round, flattened base of leavened wheat-based dough topped\n",
                            Price=25.00m,
                            ProductId=2,
                            CreateDate=DateTime.Now.AddDays(-1)
                        }
                    } 
                },
                new Category()
                {
                    Code="Category2",
                    Name="Electronics",
                    Products=new List<Product>()
                    {
                        new Product()
                        {
                            Name="Mobile Phone",
                            Desctiption=" a portable telephone that can make and receive calls over a radio frequency link ... ",
                            Price=1000.4567m,
                            ProductId=3,
                            CreateDate=DateTime.Now.AddMonths(-5)
                        },
                        new Product()
                        {
                            Name="Laptop Computer",
                            Desctiption=@"a small, portable personal computer (PC) with a ""clamshell"" form factor...",
                            Price=2000.1234m,
                            ProductId=4,
                            CreateDate=DateTime.Now.AddDays(-1)
                        }
                    }
                }
            };
        }
    }
}
