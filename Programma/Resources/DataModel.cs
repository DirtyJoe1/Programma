using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Programma.Resources
{
    public class DataModel
    {
        public short Id { get; set; }
        public string? FIO { get; set; }
        public string? Otdel { get; set; }
        public DateTime? Setup {  get; set; }
        public DateTime? Start { get; set; }
        public DateTime? End { get; set; }
        //public string ExpireStatus { get; set; }
        public string? Status { get; set; } //Подан запрос, получен, активен 
    }
}
