using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSizeZero
{
    public class DocumentModel
    {
        public int DocId { get; set; }
        public string Designation { get; set; }
        public string Name { get; set; }
        public string DirectoryPath { get; set; }
        public string FileName { get; set; }
        public byte[] FileBody { get; set; }
    }
}
