using System;
using System.Collections.Generic;
using System.Linq;

namespace DutyArranger.Source.Entities
{
    public class Soldier
    {
        private List<DateTime> _holidayDays;
        private List<DateTime> _daysOnGuard;
        private List<DateTime> _daysOnReserve;
        private int _lastDayOnDutyFromPreviousMonth;
        private string _name;

        public Soldier()
        {
            _holidayDays = new List<DateTime> { };
            _daysOnGuard = new List<DateTime> { };
            _daysOnReserve = new List<DateTime> { };
        }

        public string FirstName
        {
            get => _name;
            set => _name = value;
        }

        public int LastDayOnDutyFromPreviousMonth
        {
            get => _lastDayOnDutyFromPreviousMonth;
            set => _lastDayOnDutyFromPreviousMonth = value;
        }

        public List<DateTime> Holidays
        {
            get => _holidayDays;
            set => _holidayDays = value;
        }

        public List<DateTime> DaysOnGuard
        {
            get => _daysOnGuard;
            set => _daysOnGuard = value;
        }

        public List<DateTime> DaysOnReserve
        {
            get => _daysOnReserve;
            set => _daysOnReserve = value;
        }

        public int TimesOnDutyThisMonth(int month)
        {
            return _daysOnGuard.Count(day => day.Month == month) + _daysOnReserve.Count(day => day.Month == month);
        }

        public bool HasAlreadyBeenOnDutyThisMonth(int month)
        {
            return HasAlreadyGuardedThisMonth(month) || HasAlreadyReservedThisMonth(month);
        }

        public bool HasAlreadyGuardedThisMonth(int month)
        {
            return _daysOnGuard.Any(day => day.Month == month);
        }

        public bool HasAlreadyReservedThisMonth(int month)
        {
            return _daysOnReserve.Any(day => day.Month == month);
        }
    }
}
