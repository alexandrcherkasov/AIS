using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Helper
{
    public class Excel
    {
        private XLWorkbook wb;
        private IXLWorksheet ws;
        private List<Header> HeaderList;

        public Excel()
        {
            wb = new XLWorkbook();
            ws = wb.Worksheets.Add("Данные");
        }

        #region AIIS_MET_REP_ODPU_CIT

        /// <summary>
        /// Создание отчета AIIS_MET_REP_ODPU_CIT
        /// </summary>
        /// <param name="file">Наименование файла</param>
        /// <param name="dataTable">Данные</param>
        public void Create_AIIS_MET_REP_ODPU_CIT(string file, DataTable dataTable)
        {
            Init_AIIS_MET_REP_ODPU_CIT_HeaderList();
            InitReport(file, dataTable);
        }
        private void Init_AIIS_MET_REP_ODPU_CIT_HeaderList()
        {
            HeaderList = new List<Header>
            {
                new Header("ИСУ", new List<string>{ "Идентификатор", "Номер точки", "Тип ПУ", "Коэффициент", "Серийный СИТЭЛ", "Серийный АИИС", "Тип ПУ СИТЭЛ"}),
                new Header("Адрес", new List<string> { "Населенный пункт", "Район(обл)", "Район(гор)", "Улица", "Дома", "Квартира" }),
                new Header("СИТЭЛ", new List<string>{ "NIF", "HID", "FUH", "IS_CITEL"}),
                new Header("Текущее показание", new List<string>{ "Идентификатор", "Дата", "Показание"})
            };
        }


        #endregion

        #region AIIS_MET_RIC_REP

        /// <summary>
        /// Создание отчета AIIS_MET_RIC_REP
        /// </summary>
        /// <param name="file">Наименование файла</param>
        /// <param name="dataTable">Данные</param>
        public void Create_AIIS_MET_RIC_REP(string file, DataTable dataTable)
        {
            Init_AIIS_MET_RIC_REP_HeaderList();
            InitReport(file, dataTable);
        }

        private void Init_AIIS_MET_RIC_REP_HeaderList()
        {
            HeaderList = new List<Header>
            {
                new Header("ИСУ", new List<string> { "Номер", "Номер точки", "Серийный АИИС" }),
                new Header("Адрес", new List<string> { "Населенный пункт", "Район(обл)", "Район(гор)", "Улица", "Дома", "Квартира" }),
                new Header("Текущее показание", new List<string>{ "Идентификатор", "Дата", "Показание"}),
                new Header("Прошлое показание", new List<string>{ "Идентификатор", "Дата", "Показание"}),
                new Header("Расход между прошедшим и текущим", new List<string>{ "Разница", "Кол-во дней"}),
                new Header("УЗПУ", new List<string>{ "Потребитель", "Серийный", "Лицевой", "Дата установки", "Отделение", "Номер отделения"})
            };
        }


        #endregion

        #region AIIS_MET_OTHER_REP

        /// <summary>
        /// Создание отчета AIIS_MET_OTHER_REP
        /// </summary>
        /// <param name="file">Наименование файла</param>
        /// <param name="dataTable">Данные</param>
        public void Create_AIIS_MET_OTHER_REP(string file, DataTable dataTable)
        {
            Init_AIIS_MET_OTHER_REP_HeaderList();
            InitReport(file, dataTable);
        }

        private void Init_AIIS_MET_OTHER_REP_HeaderList()
        {
            HeaderList = new List<Header>
            {
                new Header("ИСУ", new List<string> { "Номер", "Номер точки", "Серийный АИИС" }),
                new Header("Адрес", new List<string> { "Населенный пункт", "Район(обл)", "Район(гор)", "Улица", "Дома", "Квартира" }),
                new Header("СИТЭЛ", new List<string>{ "NAF", "NIF"}),
                new Header("Текущее показание", new List<string>{ "Идентификатор", "Дата", "Показание"}),
                new Header("Прошлое показание", new List<string>{ "Идентификатор", "Дата", "Показание"}),
                new Header("Расход между прошедшим и текущим", new List<string>{ "Разница", "Кол-во дней"}),
                new Header("УЗПУ", new List<string>{ "Потребитель", "Серийный", "Лицевой", "Отделение", "Дата установки"})
            };
        }


        #endregion

        #region AIIS_MET_CITEL_REP_KEL

        /// <summary>
        /// Создание отчета AIIS_MET_CITEL_REP_KEL
        /// </summary>
        /// <param name="file">Наименование файла</param>
        /// <param name="dataTable">Данные</param>
        public void Create_AIIS_MET_CITEL_REP_KEL(string file, DataTable dataTable)
        {
            Init_AIIS_MET_OTHER_REP_HeaderList();
            InitReport(file, dataTable);
        }

        private void Init_AIIS_MET_CITEL_REP_KEL_HeaderList()
        {
            HeaderList = new List<Header>
            {
                new Header("ИСУ", new List<string> { "Номер", "Номер точки", "Серийный АИИС" }),
                new Header("Адрес", new List<string> { "Населенный пункт", "Район(обл)", "Район(гор)", "Улица", "Дома", "Квартира" }),
                new Header("СИТЭЛ", new List<string>{ "NAF", "NIF", "Лицевой", "Абонент"}),
                new Header("Текущее показание", new List<string>{ "Идентификатор", "Дата", "Показание"}),
                new Header("Прошлое показание", new List<string>{ "Идентификатор", "Дата", "Показание"}),
                new Header("Расход между прошедшим и текущим", new List<string>{ "Разница", "Кол-во дней"}),
                new Header("Начисление", new List<string>{ "Статус", "Значение", "Старое", "Новое", "Начислено"})
            };
        }


        #endregion

        #region AIIS_MET_ESKK_REP

        /// <summary>
        /// Создание отчета AIIS_MET_ESKK_REP
        /// </summary>
        /// <param name="file">Наименование файла</param>
        /// <param name="dataTable">Данные</param>
        public void Create_AIIS_MET_ESKK_REP(string file, DataTable dataTable)
        {
            Init_AIIS_MET_ESKK_REP_HeaderList();
            InitReport(file, dataTable);
        }

        private void Init_AIIS_MET_ESKK_REP_HeaderList()
        {
            HeaderList = new List<Header>
            {
                new Header("ИСУ", new List<string> { "Номер", "Номер точки", "Серийный АИИС" }),
                new Header("Адрес", new List<string> { "Населенный пункт", "Район(обл)", "Район(гор)", "Улица", "Дома", "Квартира" }),
                new Header("Текущее показание", new List<string>{ "Идентификатор", "Дата", "Показание"}),
                new Header("Прошлое показание", new List<string>{ "Идентификатор", "Дата", "Показание"}),
                new Header("Расход между прошедшим и текущим", new List<string>{ "Разница", "Кол-во дней"}),
                new Header("УЗПУ", new List<string>{ "Потребитель", "Серийный", "Лицевой", "Дата установки"})
            };
        }


        #endregion


        /// <summary>
        /// Инициализация Заголовка, данных и сохранение excel файла
        /// </summary>
        /// <param name="file"> Наименование файла</param>
        /// <param name="dataTable"> Данные для файла</param>
        private void InitReport(string file, DataTable dataTable)
        {
            CreateHeader();
            InsertDataTable(dataTable);
            SaveExcel(file);
        }

        private void InsertDataTable(DataTable dataTable)
        {
            ///Данные вносятся с 3 строки, где 3 строка, а 1 столбец
            ws.Cell(3, 1).InsertData(dataTable.AsEnumerable());
        }

        private void SaveExcel(string text)
        {
            wb.SaveAs(text + ".xlsx");
        }

        private void CreateHeader()
        {
            int HeaderCellRow = 1;
            int HeaderCellColumn = 1;
            int ChilderCellRow = 2;
            int ChilderCellColumn= 1;
            foreach(var header in HeaderList)
            {
                ws.Cell(HeaderCellRow, HeaderCellColumn).Value = header.HeaderName;
                ws.Cell(HeaderCellRow, HeaderCellColumn).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Range(HeaderCellRow, HeaderCellColumn, HeaderCellRow, HeaderCellColumn + header.MergeCount()).Merge();
                ws.Column(HeaderCellColumn).AdjustToContents();
                foreach (var childer in header.HeaderChildern)
                {
                    ws.Cell(ChilderCellRow, ChilderCellColumn).Value = childer;
                    ws.Cell(ChilderCellRow, ChilderCellColumn).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    ws.Column(ChilderCellColumn).AdjustToContents();
                    ChilderCellColumn++;
                }
                ws.Range(HeaderCellRow, HeaderCellColumn, ChilderCellRow, HeaderCellColumn + header.MergeCount()).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                HeaderCellColumn += header.MergeCount() + 1;
            }
        }
    }
}
