# Date & Time Formula

---

### DATE
**Description:**  
Returns a date value from separate year, month, and day values.

**Syntax:**  
`=DATE(year, month, day)`

**Parameters:**  
- `year`: 4-digit or 2-digit year.
- `month`: Month number (1–12). Can be >12 or <1 for overflows.
- `day`: Day number (1–31). Can be >31 or <1 for overflows.

**Example:**  
=DATE(2023, 12, 31)


Returns: 31-Dec-2023

---

### DATEVALUE
**Description:**  
Converts a date in text format to an Excel serial number.

**Syntax:**  
`=DATEVALUE(date_text)`

**Parameters:**  
- `date_text`: Text string representing a date.

**Example:**  
=DATEVALUE("04/11/2025")


Returns: Serial number for 4-Nov-2025

---

### DAY
**Description:**  
Extracts the day of the month (1–31) from a date.

**Syntax:**  
`=DAY(serial_number)`

**Parameters:**  
- `serial_number`: A valid Excel date.

**Example:**  
=DAY(DATE(2025,11,3))


Returns: 3

---

### DAYS
**Description:**  
Returns the number of days between two dates.

**Syntax:**  
`=DAYS(end_date, start_date)`

**Parameters:**  
- `end_date`: The later date.
- `start_date`: The earlier date.

**Example:**  
=DAYS("11/10/2025","10/31/2025")


Returns: 10

---

### DAYS360
**Description:**  
Calculates the number of days between two dates, based on a 360-day year.

**Syntax:**  
`=DAYS360(start_date, end_date, [method])`

**Parameters:**  
- `start_date`: Starting date.
- `end_date`: Ending date.
- `method` (optional): TRUE for European method, FALSE/default for US method.

**Example:**  
=DAYS360("01/01/2025","12/31/2025")


Returns: Number of 30-day months between dates.

---

### EDATE
**Description:**  
Returns a date a specified number of months before or after another date.

**Syntax:**  
`=EDATE(start_date, months)`

**Parameters:**  
- `start_date`: The date to start from.
- `months`: Number of months before/after.

**Example:**  
=EDATE("01/31/2025", 3)


Returns: 30-Apr-2025

---

### EOMONTH
**Description:**  
Returns the last day of the month, n months before or after specified date.

**Syntax:**  
`=EOMONTH(start_date, months)`

**Parameters:**  
- `start_date`: Start date.
- `months`: Number of months before/after.

**Example:**  
=EOMONTH("03/15/2025", 1)


Returns: 30-Apr-2025

---

### HOUR
**Description:**  
Returns the hour (0–23) from a given time.

**Syntax:**  
`=HOUR(serial_number)`

**Parameters:**  
- `serial_number`: A valid Excel time.

**Example:**  
=HOUR("15:45")


Returns: 15

---

### ISOWEEKNUM
**Description:**  
Returns the ISO week number of the year for a given date.

**Syntax:**  
`=ISOWEEKNUM(date)`

**Parameters:**  
- `date`: The date for calculation.

**Example:**  
=ISOWEEKNUM("11/10/2025")


Returns: ISO week number (example).

---

### MINUTE
**Description:**  
Returns the minute (0–59) from a given time.

**Syntax:**  
`=MINUTE(serial_number)`

**Parameters:**  
- `serial_number`: A valid Excel time.

**Example:**  
=MINUTE("15:45")


Returns: 45

---

### MONTH
**Description:**  
Returns the month (1–12) for a given date.

**Syntax:**  
`=MONTH(serial_number)`

**Parameters:**  
- `serial_number`: An Excel date.

**Example:**  
=MONTH(DATE(2025,11,10))


Returns: 11

---

### NETWORKDAYS
**Description:**  
Returns workdays between two dates, excluding weekends and optionally holidays.

**Syntax:**  
`=NETWORKDAYS(start_date, end_date, [holidays])`

**Parameters:**  
- `start_date`: Start date.
- `end_date`: End date.
- `holidays` (optional): Range of holiday dates.

**Example:**  
=NETWORKDAYS("11/01/2025","11/10/2025")


Returns: Number of workdays (excludes weekends).

---

### NETWORKDAYS.INTL
**Description:**  
Returns workdays between two dates with custom weekend days and optional holidays.

**Syntax:**  
`=NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])`

**Parameters:**  
- `start_date`: Start date.
- `end_date`: End date.
- `weekend` (optional): String or number for weekend days.
- `holidays` (optional): Range of holiday dates.

**Example:**  
=NETWORKDAYS.INTL("11/01/2025","11/10/2025",1)


Returns: Workdays with specified weekends.

---

### NOW
**Description:**  
Returns current date and time.

**Syntax:**  
`=NOW()`

**Parameters:**  
None.

**Example:**  
=NOW()


Returns: Current system date & time.

---

### SECOND
**Description:**  
Returns the second (0–59) of a given time.

**Syntax:**  
`=SECOND(serial_number)`

**Parameters:**  
- `serial_number`: Time value.

**Example:**  
=SECOND("15:45:20")


Returns: 20

---

### TIME
**Description:**  
Returns a decimal number for a specific time.

**Syntax:**  
`=TIME(hour, minute, second)`

**Parameters:**  
- `hour`: Hour (0–23).
- `minute`: Minute (0–59).
- `second`: Second (0–59).

**Example:**  
=TIME(15,45,20)


Returns: 0.656 (Excel time serial value).

---

### TIMEVALUE
**Description:**  
Converts time in text format to Excel serial number.

**Syntax:**  
`=TIMEVALUE(time_text)`

**Parameters:**  
- `time_text`: Text string representing time.

**Example:**  
=TIMEVALUE("15:45:20")


Returns: Serial number for the time.

---

### TODAY
**Description:**  
Returns the current date.

**Syntax:**  
`=TODAY()`

**Parameters:**  
None.

**Example:**  
=TODAY()


Returns: Today's date.

---

### WEEKDAY
**Description:**  
Returns number (1–7) for day of the week.

**Syntax:**  
`=WEEKDAY(serial_number, [return_type])`

**Parameters:**  
- `serial_number`: Date value.
- `return_type` (optional): Specifies week start (1=Sunday, 2=Monday, 3=Monday).

**Example:**  
=WEEKDAY("11/10/2025", 2)


Returns: 1 for Monday, 2 for Tuesday, etc.

---

### WEEKNUM
**Description:**  
Returns the week number for a given date.

**Syntax:**  
`=WEEKNUM(serial_number, [return_type])`

**Parameters:**  
- `serial_number`: Date value.
- `return_type` (optional): Specifies week starts on Sunday (1) or Monday (2).

**Example:**  
=WEEKNUM("11/10/2025")


Returns: Week number in year.

---

### WORKDAY
**Description:**  
Returns a date n workdays before/after a start date, excluding weekends & holidays.

**Syntax:**  
`=WORKDAY(start_date, days, [holidays])`

**Parameters:**  
- `start_date`: Starting date.
- `days`: Number of workdays to move (forward or backward).
- `holidays` (optional): Range of holiday dates.

**Example:**  
=WORKDAY("11/01/2025", 5)


Returns: Next 5th workday from start.

---

### WORKDAY.INTL
**Description:**  
Returns a date n workdays before/after start date with custom weekends & holidays.

**Syntax:**  
`=WORKDAY.INTL(start_date, days, [weekend], [holidays])`

**Parameters:**  
- `start_date`: Starting date.
- `days`: Number of workdays to move.
- `weekend` (optional): String/number for weekends.
- `holidays` (optional): Holiday dates.

**Example:**  
=WORKDAY.INTL("11/01/2025", 5, 1)


Returns: Next 5th workday from start, custom weekends.

---

### YEAR
**Description:**  
Returns the year from a given date.

**Syntax:**  
`=YEAR(serial_number)`

**Parameters:**  
- `serial_number`: Date value.

**Example:**  
=YEAR(DATE(2025,11,10))


Returns: 2025

---

### YEARFRAC
**Description:**  
Calculates the fraction of the year between two dates.

**Syntax:**  
`=YEARFRAC(start_date, end_date, [basis])`

**Parameters:**  
- `start_date`: Start date.
- `end_date`: End date.
- `basis` (optional): Day count basis.

**Example:**  
=YEARFRAC("01/01/2025", "11/10/2025")


Returns: Fraction of year elapsed.

---
