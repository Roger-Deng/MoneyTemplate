using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;
using MoneyTemplate.Service.FakeData;

namespace MoneyTemplate.Controllers
{
    public class HomeController : Controller
    {
        private static FakeDataSource fakeData = new FakeDataSource();

        public ActionResult Index(int page  = 1)
        {
            int curPage = page < 1 ? 1 : page;

            return View(fakeData.Data.OrderBy(x=>x.Id).ToPagedList(curPage, FakeDataSource.pageSize));
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}