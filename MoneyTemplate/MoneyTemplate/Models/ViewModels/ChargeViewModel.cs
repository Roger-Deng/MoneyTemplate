using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace MoneyTemplate.Models.ViewModels
{
    public class ChargeViewModel
    {
        public int Id { get; set; }

        public Type Category { get; set; }
  
        public string Name { get; set; }
    
        public DateTime Date { get; set; }
   
        public int Money { get; set; }

        public string Description { get; set; }
    }

    public enum Type
    {
        Expense = 1,
        Income 
    }

}