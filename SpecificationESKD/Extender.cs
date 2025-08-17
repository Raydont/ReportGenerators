using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using TFlex.Reporting;


    public static class Extender
    {
        public static  string GetAutorLastName(this IReportGenerationContext context)
        {
            PropertyInfo lastNameProperty = context.GetType().GetProperty("AuthorLastName");

            if (lastNameProperty!= null)
            {
                return (string)lastNameProperty.GetValue(context, null);
            }

            PropertyInfo autorInfoProperty = context.GetType().GetProperty("AuthorInfo");
            if (autorInfoProperty != null)
            {
                var autorInfo = autorInfoProperty.GetValue(context, null);
                if(autorInfo!= null)
                {
                    lastNameProperty = autorInfo.GetType().GetProperty("AuthorLastName");
                    if (lastNameProperty != null)
                    {
                        return (string)lastNameProperty.GetValue(autorInfo, null);
                    }
                }
            }

            return "";
        }
    }

