using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace XMLtoExcel
{
    public static class FileinfoExtension
    {
        public static string FileNameWithoutExtension(this FileInfo fileInfo)
        {
            string fullname = fileInfo.FullName;
            return fullname.Split('.')[0];
        }
    }
}
