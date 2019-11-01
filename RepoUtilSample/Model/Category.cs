using System;
using System.Collections.Generic;
using System.Text;

namespace RepoUtilSample
{
   public class Category
    {
        public string Code { get; set; }
        public string Name { get; set; }

        public List<Product> Products { get; set; }
    }
}
