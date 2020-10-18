using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace week05
{
    class Program
    {
        PortfolioEntities context = new PortfolioEntities();
        List<Tick> Ticks;
        static void Main(string[] args)
        {
            InitializeComponent();
            Ticks = context.Ticks.ToList();
            dataGridView1.DataSource = Ticks;
        }
    }
}
