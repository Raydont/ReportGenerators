using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SearchObjectsNotLinkFilesReport
{
    public class ReportItem
    {
        public string denotation; // базовое обозначение
        public string naimenovanie; // наименование
        public string Type;
        public int Level;
        public string stringPartMake = string.Empty; // строковая часть исполнения (Например: Сп - ЫК6.490.099-17 Сп)
        public string denotationWithOutStringPartMake;
        public string Author;
        public string Format;
        public List<int> variants = new List<int>(); // исполнения
        public List<int> variantsStr = new List<int>(); // исполнения строковые
        // создание элемента отчета
        public ReportItem(ReportParameters item)
        {

            naimenovanie = item.Name;
            this.Type = item.Type;
            this.Level = item.Level;
            this.Author = item.Author;
            this.Format = item.Format;
            var index = item.Denotation.LastIndexOf('-');
            if (index < 0)
            {
                denotation = item.Denotation;
                variants.Add(0);
            }
            else
            {
                //if (regexStringPartMake.Match(item.Denotation).Value.Trim() != string.Empty)
                //{
                //    denotation = item.Denotation.Substring(0, index) + " " +
                //                 regexStringPartMake.Match(item.Denotation).Value.Trim();
                //    stringPartMake = regexStringPartMake.Match(item.Denotation).Value.Trim();

                //    denotationWithOutStringPartMake = item.Denotation.Substring(0, index);
                //}
                //else
                    denotation = item.Denotation.Substring(0, index);
                try
                {
                    variants.Add(int.Parse(item.Denotation.Substring(index + 1).Trim()));
                }
                catch
                {
                    denotation = item.Denotation;
                    variants.Add(0);
                }

            }
        }

        public string Naimenovanie
        {
            get
            {
                return naimenovanie;
            }
        }



        // формируемое обозначение с учетом группировки исполнений
        public string Denotation
        {
            get
            {

                variants = variants.Distinct().OrderBy(t => t).ToList();
                //variants.Sort();

                var newDenotation = string.Empty;
                if (stringPartMake == string.Empty)
                {
                    newDenotation = denotation;
                    if (variants.Count == 2)
                    {
                        if (variants[1] == 1)
                            newDenotation += ", ";
                    }
                }
                else
                {
                    newDenotation = denotationWithOutStringPartMake;

                    if (variants[0] == 0)
                    {
                        newDenotation = denotation;

                        if (variants.Count == 2)
                        {
                            if (variants[1] == 1)
                                newDenotation += ", ";
                        }

                    }
                }
                if (variants.Count == 1)
                {
                    if (variants[0] != 0)
                    {
                        if (stringPartMake == string.Empty)
                            newDenotation += "-" + VariantToString(variants[0]);
                        else
                            newDenotation += "-" + VariantToString(variants[0]) + " " + stringPartMake;
                    }
                }
                else
                {
                    // создание цепочек исполнений
                    var variantChains = new List<List<int>>();
                    var lastChain = new List<int>();
                    lastChain.Add(variants[0]);
                    variantChains.Add(lastChain);

                    for (int i = 1; i < variants.Count; i++)
                    {

                        if (variants[i] - lastChain[lastChain.Count - 1] == 0)
                        {

                            continue;
                        }

                        if (variants[i] - lastChain[lastChain.Count - 1] == 1)
                        {
                            lastChain.Add(variants[i]);

                        }
                        else
                        {
                            lastChain = new List<int>();
                            lastChain.Add(variants[i]);
                            variantChains.Add(lastChain);
                        }
                    }

                    // формирование строки обозначения с учетом цепочек исполнений
                    //if (variants.Count>1)
                    //    MessageBox.Show(string.Join(", ", ", "+variantChains.Select(ChainToString).ToArray()));

                    if (stringPartMake == string.Empty)
                        newDenotation += string.Join(", ", variantChains.Select(ChainToString).ToArray());
                    else
                        newDenotation += string.Join(", ", variantChains.Select(ChainToStringWithPartMake).ToArray());

                }

                return newDenotation;

            }
        }




        // преобразование цепочки вариантов в строковое представление
        private string ChainToString(List<int> chain)
        {

            if (chain.Count == 1)
            {
                if (chain[0] != 0)
                    return "-" + VariantToString(chain[0]);
                else return string.Empty;
            }
            if (chain.Count == 2)
            {
                if (chain[0] != 0)
                {
                    return "-" + VariantToString(chain[0]) + ", -" + VariantToString(chain[1]);
                }
                else
                {
                    return "-" + VariantToString(chain[1]);
                }
            }
            if (chain.Count == 3)
            {
                if (chain[0] != 0)
                {
                    return "-" + VariantToString(chain[0]) + "...-" + VariantToString(chain[chain.Count - 1]);
                }
                else
                {
                    return "-" + VariantToString(chain[1]) + ", -" + VariantToString(chain[2]);
                }
            }


            if (chain[0] != 0)
            {
                return "-" + VariantToString(chain[0]) + "...-" + VariantToString(chain[chain.Count - 1]);
            }
            else
            {
                return "-" + VariantToString(chain[1]) + "...-" + VariantToString(chain[chain.Count - 1]);
            }
        }

        private string ChainToStringWithPartMake(List<int> chain)
        {

            if (chain.Count == 1)
            {
                if (chain[0] != 0)
                    return "-" + VariantToString(chain[0]) + " " + stringPartMake;
                else return string.Empty;
            }
            if (chain.Count == 2)
            {
                if (chain[0] != 0)
                {
                    return "-" + VariantToString(chain[0]) + " " + stringPartMake + ", -" + VariantToString(chain[1]) + " " + stringPartMake;
                }
                else
                {
                    return "-" + VariantToString(chain[1]) + " " + stringPartMake;
                }
            }
            if (chain.Count == 3)
            {
                if (chain[0] != 0)
                {
                    return "-" + VariantToString(chain[0]) + " " + stringPartMake + "...-" + VariantToString(chain[chain.Count - 1]) + " " + stringPartMake;
                }
                else
                {
                    return "-" + VariantToString(chain[1]) + " " + stringPartMake + ", -" + VariantToString(chain[2]) + " " + stringPartMake;
                }
            }


            if (chain[0] != 0)
            {
                return "-" + VariantToString(chain[0]) + " " + stringPartMake + "...-" + VariantToString(chain[chain.Count - 1]) + " " + stringPartMake;
            }
            else
            {
                return "-" + VariantToString(chain[1]) + " " + stringPartMake + "...-" + VariantToString(chain[chain.Count - 1]) + " " + stringPartMake;
            }
        }


        // преобразование варианта в строкове представление
        private string VariantToString(int variant)
        {
            if (variant > 99)
            {
                return variant.ToString("000");
            }
            return variant.ToString("00");
        }

        // проверка нового элемента на соответствие текущему (является ли элемент исполнением)
        public bool IsValid(ReportParameters item)
        {
            var index = item.Denotation.LastIndexOf('-');
            var itemDenotationBase = ""; // обозначение без части исполнения
          
            if (index < 0)
            {
                itemDenotationBase = item.Denotation;
            }
            else
            {
                itemDenotationBase = item.Denotation.Substring(0, index);
            }

            if ((itemDenotationBase != denotation)) // обозначения не соответствуют
            {

                return false;
            }


            if (index >= 0)
            {
                var variantString = item.Denotation.Substring(index + 1).Trim();
                //if (variantString.Length == 1)
                //{
                //    MessageBox.Show(denotation);
                //    return false; // если одна цифра исполнения 
                //}
                for (int i = 0; i < variantString.Length; i++) // если часть исполнения содержит не цифры
                {
                    if (!char.IsDigit(variantString[i]))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        // функция добавления исполнения
        public void Add(ReportParameters item)
        {
            var index = item.Denotation.LastIndexOf('-');
            if (index < 0)
            {
                variants.Add(0);
            }
            else
            {
                try
                {
                    variants.Add(int.Parse(item.Denotation.Substring(index + 1)));
                }
                catch
                {
                    denotation = item.Denotation;
                    // variants.Add(999999);
                   // variants.Add(int.Parse(regexNumberMake.Match(item.Denotation).Value));
                    // MessageBox.Show("Не удалось добавить вариант " + item.Denotation + " к " + denotation + "!");
                }
            }
        }
    }
}
