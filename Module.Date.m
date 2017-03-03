// Copyright (c) Microsoft Corporation.  All rights reserved.
    
section Date;

shared Date.IsInPreviousDay = (dateTime) as nullable logical =>
    Date.IsInPreviousNDays(dateTime, 1);

shared Date.IsInPreviousNDays = (dateTime, days as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.AddDays(DateTime.FixedLocalNow(), -days)) and
        Date.From(dateTime) < Date.From(DateTime.FixedLocalNow());
        
shared Date.IsInCurrentDay = (dateTime) as nullable logical =>
    Date.From(dateTime) >= Date.From(DateTime.FixedLocalNow()) and
        Date.From(dateTime) < Date.From(Date.AddDays(DateTime.FixedLocalNow(), 1));

shared Date.IsInNextDay = (dateTime) as nullable logical =>
    Date.IsInNextNDays(dateTime, 1);

shared Date.IsInNextNDays = (dateTime, days as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.AddDays(DateTime.FixedLocalNow(), 1)) and 
        Date.From(dateTime) < Date.From(Date.AddDays(DateTime.FixedLocalNow(), days + 1));
            
shared Date.IsInPreviousWeek = (dateTime) as nullable logical =>
    Date.IsInPreviousNWeeks(dateTime, 1);

shared Date.IsInPreviousNWeeks = (dateTime, weeks as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.StartOfWeek(Date.AddWeeks(DateTime.FixedLocalNow(), -weeks))) and
        Date.From(dateTime) < Date.From(Date.StartOfWeek(DateTime.FixedLocalNow()));

shared Date.IsInCurrentWeek = (dateTime) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.StartOfWeek(DateTime.FixedLocalNow())) and
        Date.From(dateTime) < Date.From(Date.AddDays(Date.EndOfWeek(DateTime.FixedLocalNow()), 1));

shared Date.IsInNextWeek = (dateTime) as nullable logical =>
    Date.IsInNextNWeeks(dateTime, 1);

shared Date.IsInNextNWeeks = (dateTime, weeks as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.AddDays(Date.EndOfWeek(DateTime.FixedLocalNow()), 1)) and
        Date.From(dateTime) < Date.From(Date.AddDays(Date.EndOfWeek(Date.AddWeeks(DateTime.FixedLocalNow(), weeks)), 1));

shared Date.IsInPreviousMonth = (dateTime) as nullable logical =>
    Date.IsInPreviousNMonths(dateTime, 1);

shared Date.IsInPreviousNMonths = (dateTime, months as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.StartOfMonth(Date.AddMonths(DateTime.FixedLocalNow(), -months))) and
        Date.From(dateTime) < Date.From(Date.StartOfMonth(DateTime.FixedLocalNow()));

shared Date.IsInCurrentMonth = (dateTime) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.StartOfMonth(DateTime.FixedLocalNow())) and
        Date.From(dateTime) < Date.From(Date.AddDays(Date.EndOfMonth(DateTime.FixedLocalNow()), 1));

shared Date.IsInNextMonth = (dateTime) as nullable logical =>
    Date.IsInNextNMonths(dateTime, 1);

shared Date.IsInNextNMonths = (dateTime, months as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.AddDays(Date.EndOfMonth(DateTime.FixedLocalNow()), 1)) and
        Date.From(dateTime) < Date.From(Date.AddDays(Date.EndOfMonth(Date.AddMonths(DateTime.FixedLocalNow(), months)), 1));

shared Date.IsInPreviousQuarter = (dateTime) as nullable logical =>
    Date.IsInPreviousNQuarters(dateTime, 1);

shared Date.IsInPreviousNQuarters = (dateTime, quarters as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.StartOfQuarter(Date.AddQuarters(DateTime.FixedLocalNow(), -quarters))) and
        Date.From(dateTime) < Date.From(Date.StartOfQuarter(DateTime.FixedLocalNow()));

shared Date.IsInCurrentQuarter = (dateTime) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.StartOfQuarter(DateTime.FixedLocalNow())) and
        Date.From(dateTime) < Date.From(Date.AddDays(Date.EndOfQuarter(DateTime.FixedLocalNow()), 1));

shared Date.IsInNextQuarter = (dateTime) as nullable logical =>
    Date.IsInNextNQuarters(dateTime, 1);

shared Date.IsInNextNQuarters = (dateTime, quarters as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.AddDays(Date.EndOfQuarter(DateTime.FixedLocalNow()), 1)) and
        Date.From(dateTime) < Date.From(Date.AddDays(Date.EndOfQuarter(Date.AddQuarters(DateTime.FixedLocalNow(), quarters)), 1));

shared Date.IsInPreviousYear = (dateTime) as nullable logical =>
    Date.IsInPreviousNYears(dateTime, 1);

shared Date.IsInPreviousNYears = (dateTime, years as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.StartOfYear(Date.AddYears(DateTime.FixedLocalNow(), -years))) and
        Date.From(dateTime) < Date.From(Date.StartOfYear(DateTime.FixedLocalNow()));

shared Date.IsInCurrentYear = (dateTime) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.StartOfYear(DateTime.FixedLocalNow())) and
        Date.From(dateTime) < Date.From(Date.AddDays(Date.EndOfYear(DateTime.FixedLocalNow()), 1));

shared Date.IsInNextYear = (dateTime) as nullable logical =>
    Date.IsInNextNYears(dateTime, 1);

shared Date.IsInNextNYears = (dateTime, years as number) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.AddDays(Date.EndOfYear(DateTime.FixedLocalNow()), 1)) and
        Date.From(dateTime) < Date.From(Date.AddDays(Date.EndOfYear(Date.AddYears(DateTime.FixedLocalNow(), years)), 1));

shared Date.IsInYearToDate = (dateTime) as nullable logical =>
    Date.From(dateTime) >= Date.From(Date.StartOfYear(DateTime.FixedLocalNow())) and
        Date.From(dateTime) < Date.From(Date.AddDays(Date.EndOfDay(DateTime.FixedLocalNow()), 1));

shared DateTime.IsInPreviousSecond = (dateTime) as nullable logical =>
    DateTime.IsInPreviousNSeconds(dateTime, 1);

shared DateTime.IsInPreviousNSeconds = (dateTime, seconds as number) as nullable logical =>
    DateTime.From(dateTime) >= StartOfSecond - #duration(0, 0, 0, seconds) and
        DateTime.From(dateTime) < StartOfSecond;

shared DateTime.IsInNextSecond = (dateTime) as nullable logical =>
    DateTime.IsInNextNSeconds(dateTime, 1);

shared DateTime.IsInNextNSeconds = (dateTime, seconds as number) as nullable logical =>
    DateTime.From(dateTime) >= StartOfSecond + #duration(0, 0, 0, 1) and
	    DateTime.From(dateTime) < StartOfSecond + #duration(0, 0, 0, seconds + 1);

shared DateTime.IsInCurrentSecond = (dateTime) as nullable logical =>
    DateTime.From(dateTime) >= StartOfSecond and
	    DateTime.From(dateTime) < StartOfSecond + #duration(0, 0, 0, 1);

shared DateTime.IsInPreviousMinute = (dateTime) as nullable logical =>
    DateTime.IsInPreviousNMinutes(dateTime, 1);

shared DateTime.IsInPreviousNMinutes = (dateTime, minutes as number) as nullable logical =>
    DateTime.From(dateTime) >= StartOfMinute - #duration(0, 0, minutes, 0) and
        DateTime.From(dateTime) < StartOfMinute;

shared DateTime.IsInNextMinute = (dateTime) as nullable logical =>
    DateTime.IsInNextNMinutes(dateTime, 1);

shared DateTime.IsInNextNMinutes = (dateTime, minutes as number) as nullable logical =>
    DateTime.From(dateTime) >= StartOfMinute + #duration(0, 0, 1, 0) and 
	    DateTime.From(dateTime) < StartOfMinute + #duration(0, 0, minutes + 1, 0);

shared DateTime.IsInCurrentMinute = (dateTime) as nullable logical =>
    DateTime.From(dateTime) >= StartOfMinute and
	    DateTime.From(dateTime) < StartOfMinute + #duration(0, 0, 1, 0);

shared DateTime.IsInPreviousHour = (dateTime) as nullable logical =>
    DateTime.IsInPreviousNHours(dateTime, 1);

shared DateTime.IsInPreviousNHours = (dateTime, hours as number) as nullable logical =>
    DateTime.From(dateTime) >= StartOfHour - #duration(0, hours, 0, 0) and
        DateTime.From(dateTime) < StartOfHour;

shared DateTime.IsInNextHour = (dateTime) as nullable logical =>
    DateTime.IsInNextNHours(dateTime, 1);

shared DateTime.IsInNextNHours = (dateTime, hours as number) as nullable logical =>
    DateTime.From(dateTime) >= StartOfHour + #duration(0, 1, 0, 0) and 
	    DateTime.From(dateTime) < StartOfHour + #duration(0, hours + 1, 0, 0);

shared DateTime.IsInCurrentHour = (dateTime) as nullable logical =>
    DateTime.From(dateTime) >= StartOfHour and
	    DateTime.From(dateTime) < StartOfHour + #duration(0, 1, 0, 0);

shared Date.MonthName = (date, optional culture as text) as nullable text => 
    DateTimeZone.ToText(DateTimeZone.From(date), "MMMM", culture);

shared Date.DayOfWeekName = (date, optional culture as text) as nullable text => 
    DateTimeZone.ToText(DateTimeZone.From(date), "dddd", culture);

StartOfHour = DateTime.From(Time.StartOfHour(DateTime.FixedLocalNow()));

StartOfMinute = DateTime.From(Time.StartOfHour(DateTime.FixedLocalNow()) + #duration(0, 0, Time.Minute(DateTime.FixedLocalNow()), 0));

StartOfSecond = DateTime.From(Time.StartOfHour(DateTime.FixedLocalNow()) + #duration(0, 0, Time.Minute(DateTime.FixedLocalNow()), Number.RoundDown(Time.Second(DateTime.FixedLocalNow()))));
    
