using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MoneyTemplate.Models.ViewModels;

namespace MoneyTemplate.Models.FakeData
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
            _data = new List<ChargeViewModel> {
                new ChargeViewModel{ Id=1, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 50  },
                new ChargeViewModel{ Id=2, Category = ViewModels.Type.Income, Name = "Dinner",  Date = System.DateTime.Now, Money = 30  },
                new ChargeViewModel{ Id=3, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=4, Category = ViewModels.Type.Income, Name = "Lunch",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=5, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=6, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 3000  },
                new ChargeViewModel{ Id=7, Category = ViewModels.Type.Income, Name = "Dinner",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=8, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 190  },
                new ChargeViewModel{ Id=9, Category = ViewModels.Type.Expense, Name = "Dinner",  Date = System.DateTime.Now, Money = 400  },
                new ChargeViewModel{ Id=10, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 320  },
                new ChargeViewModel{ Id=11, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 500  },
                new ChargeViewModel{ Id=12, Category = ViewModels.Type.Income, Name = "Dinner",  Date = System.DateTime.Now, Money = 300  },
                new ChargeViewModel{ Id=13, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 350  },
                new ChargeViewModel{ Id=14, Category = ViewModels.Type.Income, Name = "Lunch",  Date = System.DateTime.Now, Money = 650  },
                new ChargeViewModel{ Id=15, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=16, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 3000  },
                new ChargeViewModel{ Id=17, Category = ViewModels.Type.Income, Name = "Dinner",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=18, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 200  },
                new ChargeViewModel{ Id=19, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 100  },
                new ChargeViewModel{ Id=20, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 280  },
                new ChargeViewModel{ Id=21, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 600  },
                new ChargeViewModel{ Id=22, Category = ViewModels.Type.Income, Name = "Dinner",  Date = System.DateTime.Now, Money = 800  },
                new ChargeViewModel{ Id=23, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 450  },
                new ChargeViewModel{ Id=24, Category = ViewModels.Type.Income, Name = "Lunch",  Date = System.DateTime.Now, Money = 750  },
                new ChargeViewModel{ Id=25, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 250  },
                new ChargeViewModel{ Id=26, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 290  },
                new ChargeViewModel{ Id=27, Category = ViewModels.Type.Income, Name = "Dinner",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=28, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 5000  },
                new ChargeViewModel{ Id=29, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 9500  },
                new ChargeViewModel{ Id=30, Category = ViewModels.Type.Expense, Name = "Dinner",  Date = System.DateTime.Now, Money = 6200  },
                new ChargeViewModel{ Id=31, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 500  },
                new ChargeViewModel{ Id=32, Category = ViewModels.Type.Income, Name = "Dinner",  Date = System.DateTime.Now, Money = 300  },
                new ChargeViewModel{ Id=33, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 350  },
                new ChargeViewModel{ Id=34, Category = ViewModels.Type.Income, Name = "Lunch",  Date = System.DateTime.Now, Money = 140  },
                new ChargeViewModel{ Id=35, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=36, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 3500  },
                new ChargeViewModel{ Id=37, Category = ViewModels.Type.Income, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=38, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 3000  },
                new ChargeViewModel{ Id=39, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 7500  },
                new ChargeViewModel{ Id=40, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 9200  },
                new ChargeViewModel{ Id=41, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 1950  },
                new ChargeViewModel{ Id=42, Category = ViewModels.Type.Income, Name = "Dinner",  Date = System.DateTime.Now, Money = 4370  },
                new ChargeViewModel{ Id=43, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=44, Category = ViewModels.Type.Income, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=45, Category = ViewModels.Type.Expense, Name = "Fastbreak",  Date = System.DateTime.Now, Money = 150  },
                new ChargeViewModel{ Id=46, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 9331  },
                new ChargeViewModel{ Id=47, Category = ViewModels.Type.Income, Name = "Dinner",  Date = System.DateTime.Now, Money = 1670  },
                new ChargeViewModel{ Id=48, Category = ViewModels.Type.Expense, Name = "Lunch",  Date = System.DateTime.Now, Money = 2700  },
                new ChargeViewModel{ Id=49, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 3200  },
                new ChargeViewModel{ Id=50, Category = ViewModels.Type.Expense, Name = "Traffic-Fee",  Date = System.DateTime.Now, Money = 6500  }
            };
        }

    }
}