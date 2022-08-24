using System;
using System.Collections.ObjectModel;

namespace DutyArranger.Source.Entities
{
    public class SelectedYear
    {
        public int Year;
        public SelectedMonth SelectedMonth = new SelectedMonth();
    }

    public class SelectedMonth
    {
        public int Month;
        public Collection<SelectedDay> SelectedDays = new Collection<SelectedDay>();
    }

    public class SelectedDay
    {
        public DateTime Date;
        public int Day;
        public Soldier Guard;
        public Soldier Reservee;
    }
}
