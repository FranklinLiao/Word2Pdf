using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word2Pdf
{
    class FileToChange
    {
        private String fileName;
        private int fileSize;
        public string FileName { get => fileName; set => fileName = value; }
        public int FileSize { get => fileSize; set => fileSize = value; }
    }
}
