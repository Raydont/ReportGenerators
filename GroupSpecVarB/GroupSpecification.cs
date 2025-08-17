using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupSpecVarB
{
    /// <summary>
    /// Класс групповой спецификации
    /// </summary>
    internal class GroupSpecification
    {
        private List<ReportRow> _resultReportRows = new List<ReportRow>();

        /// <summary>
        /// Инициализация групповой спецификации с помощью списка спецификаций
        /// </summary>
        /// <param name="specifications">набор спецификаций</param>
      /*  public GroupSpecification(SortedDictionary<int, Specification> specifications)
        {
            if (specifications == null)
                throw new NullReferenceException();

            if (specifications.Count == 0 || !specifications.ContainsKey(0))
                throw new InvalidDataException("Не хватает данных для построения групповой спецификации!");
//
           // foreach (Specification currentSpec in specifications.Values)
            //    _resultReportRows.AddRange(currentSpec.SpecificationRows);

            _resultReportRows.Sort();
            MergeDublicates();
        }*/
        /// <summary>
        /// Определяем повторяющиеся записи
        /// </summary>
        public void MergeDublicates()
        {
            for (int masterIndex = 0; masterIndex < _resultReportRows.Count; masterIndex++)
            {
                ReportRow MasterRow = _resultReportRows[masterIndex];

                for (int slaveIndex = masterIndex + 1; slaveIndex < _resultReportRows.Count; slaveIndex++)
                {
                    ReportRow SlaveRow = _resultReportRows[slaveIndex];

                    if (SlaveRow.LookLike(MasterRow))
                    {
                        MasterRow.SameRows.Add(SlaveRow);
                        _resultReportRows.RemoveAt(slaveIndex);
                        slaveIndex--;
                    }
                }
            }
        }
        /// <summary>
        /// Возвращает максимальный номер исполнения групповой спецификации
        /// </summary>
        /// <returns>максимальный номер исполнения</returns>
        public int GetMaxVariantNumber()
        {
            int MaxResult = 0;

            foreach (ReportRow row in _resultReportRows)
            {
                if (row.ParentDocVariantNumber > MaxResult)
                    MaxResult = row.ParentDocVariantNumber;

                foreach (ReportRow variantRow in row.SameRows)
                    if (variantRow.ParentDocVariantNumber > MaxResult)
                        MaxResult = variantRow.ParentDocVariantNumber;
            }

            return MaxResult;
        }
        /// <summary>
        /// Возвращает кол-во наборов исполнений для вставки по 10 (т.н. кол-во секций)
        /// </summary>
        /// <returns>кол-во секций</returns>
        public int GetSectionsCount()
        {
            int MaxVariantNumber = GetMaxVariantNumber();
            return (int)(MaxVariantNumber / 10) + 1;
        }
        /// <summary>
        /// Возвращает список документов для указанного номера секции
        /// </summary>
        /// <param name="sectionNumber">номер секции</param>
        /// <returns>списко строк для указанного номера секции</returns>
        public List<List<ReportRow>> GetSectionRowList(int sectionNumber)
        {
            List<List<ReportRow>> ResultRows = new List<List<ReportRow>>();
            int SectionsCount = GetSectionsCount();

            if (sectionNumber > SectionsCount)
                return ResultRows;

            int MaxVarNumber = sectionNumber * 10 + 9;
            int MinVarNumber = sectionNumber * 10;

            foreach (ReportRow currentRow in _resultReportRows)
            {
                List<ReportRow> RowsList = new List<ReportRow>();

                if (currentRow.ParentDocVariantNumber >= MinVarNumber &&
                    currentRow.ParentDocVariantNumber <= MaxVarNumber)
                    RowsList.Add(currentRow);

                foreach (ReportRow sameRows in currentRow.SameRows)
                {
                    if (sameRows.ParentDocVariantNumber >= MinVarNumber &&
                        sameRows.ParentDocVariantNumber <= MaxVarNumber)
                        RowsList.Add(sameRows);
                }

                if (RowsList.Count > 0)
                    ResultRows.Add(RowsList);
            }

            return ResultRows;
        }
    }
}
