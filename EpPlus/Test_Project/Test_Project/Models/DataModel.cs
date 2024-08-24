using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Test_Project.Models
{
    public class proDataModel
    {
       
            [Key]
            public int Id { get; set; }
            public string Column1 { get; set; }
            public string Column2 { get; set; }
            // سایر ستون‌ها بر اساس ساختار اکسل شما
        

    }
   

}
