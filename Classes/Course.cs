using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnrolmentTool.Classes
{
    class Course
    {
        public float Tution { get; set; }
        public float Percent { get; set; }
        public List<float> Payments { get; set; }
        public Course()
        {
            Payments = new List<float>();
        }
    }
}
