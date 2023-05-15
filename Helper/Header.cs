using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Helper
{
    public class Header
    {
        public string HeaderName;
        public List<string> HeaderChildern;
        /// <summary>
        /// Создание элемента шапки для Excel документа
        /// </summary>
        /// <param name="HeaderName">Основной заголовочный текст</param>
        /// <param name="HeaderChildren">Дочерние элементы под заголовком</param>
        public Header(string HeaderName, List<string> HeaderChildren)
        {
            this.HeaderName = HeaderName;
            this.HeaderChildern = new List<string>(HeaderChildren);
        }

        public int MergeCount()
        {
            return HeaderChildern.Count() - 1;
        }
    }
}
