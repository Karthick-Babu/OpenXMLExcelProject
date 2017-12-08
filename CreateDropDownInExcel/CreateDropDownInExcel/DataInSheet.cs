using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateDropDownInExcel
{
   public class DataInSheet
    {
        public string firstRow { get; set; }
        public string secondRow { get; set; }
        public string thirdRow { get; set; }
        public string fourthRow { get; set; }

        public static List<DataInSheet> GetDataOfSheet1()
        {
            List<DataInSheet> dataForSheet = new List<DataInSheet>
                                      {
                                             new DataInSheet
                                             {
                                                 firstRow = "Name",
                                                 secondRow = "Column1",
                                                 thirdRow = "Column2",
                                                 fourthRow = "Column3"
                                             }                                             
                                         };
            return dataForSheet;
        }

        public static List<DataInSheet> GetDataOfSheet2()
        {
            List<DataInSheet> dataForSecondSheet = new List<DataInSheet>
                                         {
                                             new DataInSheet
                                             {
                                                 firstRow = "Name1"
                                                
                                             },
                                             new DataInSheet
                                             {
                                                 firstRow = "Name2"
                                               
                                             },
                                             new DataInSheet
                                             {
                                                 firstRow = "Name3"
                                               
                                             },
                                              new DataInSheet
                                             {
                                                  firstRow = "Name4"
                                                 
                                             },
                                               new DataInSheet
                                             {
                                                 firstRow = "Name5"
                                               
                                             }
                                            
                                         };
            return dataForSecondSheet;
        }



    }


}
