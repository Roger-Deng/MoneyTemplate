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

        public string Category { get; set; }
  
        public string Name { get; set; }
    
        public DateTime Date { get; set; }
   
        public decimal Money { get; set; }

        public string Description { get; set; }
    }

    public enum Type
    {
        Expense = 1,
        Income 
    }

}