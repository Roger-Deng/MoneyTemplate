using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MoneyTemplate.Models.ViewModels;

namespace MoneyTemplate.Service.FakeData
{
    public class FakeDataSource
    {
        public static int pageSize = 10 ;

        private IList<ChargeViewModel> _data;

        public IList<ChargeViewModel> Data
        {
            get
            {
                if (_data == null) {
                    CreateData();
                }
                return _data;
            }
        }

        private void CreateData()
        {
            Random rand = new Random();
            //IList<string> item = new List<string> { "Lunch", "Dinner", "Fastbreak", "Traffic-Fee" };
            IList<string> type = new List<string> { "支出", "收入", "支出", "收入" };
            _data = new List<ChargeViewModel>(); 

            for (int i = 1; i <= 50; i++) {
               _data.Add(new ChargeViewModel { Id = i, Category = type[rand.Next(3)], Name = "", Date = System.DateTime.Now, Money = (decimal)rand.Next(15000) });
            }
        }

    }
}