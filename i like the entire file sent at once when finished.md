<img src="https://r2cdn.perplexity.ai/pplx-full-logo-primary-dark%402x.png" style="height:64px;margin-right:32px"/>

# i like the entire file sent at once when finished

I will now generate the complete Markdown file for all listed **Financial Formula** functions, with each entry including:

- Description
- Syntax
- Parameters explained
- Example

***

```markdown
# Excel Functions Reference

---

## Financial Formula

---

### ACCRINT

**Description:**  
Calculates the accrued interest for a security that pays periodic interest.

**Syntax:**  
`=ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis])`

**Parameters:**  
- `issue`: The issue date of the security (start date).
- `first_interest`: Date of first interest payment.
- `settlement`: Date you purchase the security.
- `rate`: Annual coupon rate (e.g., 8% as 0.08).
- `par`: Par value of the security (usually 1000).
- `frequency`: Number of coupon payments per year (1=annual, 2=semiannual, 4=quarterly).
- `basis` (optional): Day count basis (0=US 30/360, 1=Actual/Actual, 2=Actual/360, 3=Actual/365, 4=European 30/360).

**Example:**  
```

=ACCRINT("01-01-2020","01-07-2020","01-04-2020",0.08,1000,2)

```
Returns: Accrued Interest Value

---

### ACCRINTM

**Description:**  
Returns the accrued interest for a security that pays interest at maturity.

**Syntax:**  
`=ACCRINTM(issue, settlement, rate, par, [basis])`

**Parameters:**  
- `issue`: Issue date of the security.
- `settlement`: Maturity date.
- `rate`: Annual coupon rate (as a decimal).
- `par`: Par value.
- `basis` (optional): Day count basis.

**Example:**  
```

=ACCRINTM("01-01-2020", "01-01-2022", 0.08, 1000)

```
Returns: Accrued interest at maturity.

---

### AMORLINC

**Description:**  
Returns the linear depreciation for each accounting period for an asset.

**Syntax:**  
`=AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis])`

**Parameters:**  
- `cost`: Initial cost of the asset.
- `date_purchased`: Date the asset was purchased.
- `first_period`: End date of the first period.
- `salvage`: Value at end of depreciation.
- `period`: Period for which you want depreciation.
- `rate`: Depreciation rate.
- `basis` (optional): Day count basis.

**Example:**  
```

=AMORLINC(10000, "01-01-2020", "12-31-2020", 1000, 1, 0.2)

```
Returns: Linear depreciation for year 1.

---

### COUPDAYSBS

**Description:**  
Returns the number of days from the beginning of the coupon period to the settlement date.

**Syntax:**  
`=COUPDAYSBS(settlement, maturity, frequency, [basis])`

**Parameters:**  
- `settlement`: Date bond is purchased.
- `maturity`: Maturity date.
- `frequency`: Number of coupon payments per year.
- `basis` (optional): Day count basis.

**Example:**  
```

=COUPDAYSBS("01-04-2020", "01-01-2025", 2)

```
Returns: Days from start of coupon to settlement.

---

### COUPDAYS

**Description:**  
Returns the number of days in the coupon period that contains the settlement date.

**Syntax:**  
`=COUPDAYS(settlement, maturity, frequency, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: End date.
- `frequency`: Payments per year.
- `basis` (optional): Day count basis.

**Example:**  
```

=COUPDAYS("01-04-2020", "01-01-2025", 2)

```
Returns: Coupon period days.

---

### COUPDAYSNC

**Description:**  
Returns the number of days from the settlement date to the next coupon date.

**Syntax:**  
`=COUPDAYSNC(settlement, maturity, frequency, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: End date.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=COUPDAYSNC("01-04-2020", "01-01-2025", 2)

```
Returns: Days to next coupon.

---

### COUPNCD

**Description:**  
Returns the next coupon date after the settlement date.

**Syntax:**  
`=COUPNCD(settlement, maturity, frequency, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: Maturity.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=COUPNCD("01-04-2020", "01-01-2025", 2)

```
Returns: Next coupon date.

---

### COUPNUM

**Description:**  
Returns the number of coupons payable between the settlement date and maturity date.

**Syntax:**  
`=COUPNUM(settlement, maturity, frequency, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: Maturity.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=COUPNUM("01-04-2020", "01-01-2025", 2)

```
Returns: Number of coupons remaining.

---

### COUPPCD

**Description:**  
Returns the previous coupon date before the settlement date.

**Syntax:**  
`=COUPPCD(settlement, maturity, frequency, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: Maturity.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=COUPPCD("01-04-2020", "01-01-2025", 2)

```
Returns: Previous coupon date before settlement.

---

### CUMIPMT

**Description:**  
Returns the cumulative interest paid between two periods.

**Syntax:**  
`=CUMIPMT(rate, nper, pv, start_period, end_period, type)`

**Parameters:**  
- `rate`: Interest rate per period.
- `nper`: Total number of payments.
- `pv`: Present value/principal.
- `start_period`: First payment period in calculation.
- `end_period`: Last payment period in calculation.
- `type`: 0 for payments at end, 1 at beginning.

**Example:**  
```

=CUMIPMT(0.05/12, 36, 10000, 1, 12, 0)

```
Returns: Total interest paid in year 1.

---

### CUMPRINC

**Description:**  
Returns the cumulative principal paid between two periods.

**Syntax:**  
`=CUMPRINC(rate, nper, pv, start_period, end_period, type)`

**Parameters:**  
- `rate`: Interest rate per period.
- `nper`: Number of payments.
- `pv`: Present value/principal.
- `start_period`: First period in calculation.
- `end_period`: Last period in calculation.
- `type`: 0 for payments at end, 1 at beginning.

**Example:**  
```

=CUMPRINC(0.05/12, 36, 10000, 1, 12, 0)

```
Returns: Principal paid in year 1.

---

### DB

**Description:**  
Returns depreciation of an asset for period using fixed-declining balance method.

**Syntax:**  
`=DB(cost, salvage, life, period, [month])`

**Parameters:**  
- `cost`: Initial cost.
- `salvage`: Value after useful life.
- `life`: Number of periods.
- `period`: The period for which depreciation is calculated.
- `month` (optional): Number of months in the first year.

**Example:**  
```

=DB(10000, 1000, 5, 1)

```
Returns: Depreciation for period 1.

---

### DDB

**Description:**  
Returns double-declining balance depreciation for a period.

**Syntax:**  
`=DDB(cost, salvage, life, period, [factor])`

**Parameters:**  
- `cost`: Initial cost.
- `salvage`: Value after useful life.
- `life`: Number of periods.
- `period`: Period of calculation.
- `factor` (optional): Rate at which balance declines (default 2).

**Example:**  
```

=DDB(10000, 1000, 5, 1)

```
Returns: Depreciation using DDB.

---

### DISC

**Description:**  
Returns the discount rate for a security.

**Syntax:**  
`=DISC(settlement, maturity, pr, redemption, [basis])`

**Parameters:**  
- `settlement`: Date purchased.
- `maturity`: Maturity date.
- `pr`: Price per $100 face value.
- `redemption`: Redemption value per $100 face.
- `basis` (optional): Day count basis.

**Example:**  
```

=DISC("01-04-2020", "01-01-2025", 95, 100)

```
Returns: Discount rate.

---

### DOLLARDE

**Description:**  
Converts a dollar price in fractional format to decimal format.

**Syntax:**  
`=DOLLARDE(fractional_dollar, fraction)`

**Parameters:**  
- `fractional_dollar`: Dollar price in fractional.
- `fraction`: Denominator in the fraction.

**Example:**  
```

=DOLLARDE(1.02, 16)

```
Returns: $1.125

---

### DOLLARFR

**Description:**  
Converts a dollar price in decimal format to fractional format.

**Syntax:**  
`=DOLLARFR(decimal_dollar, fraction)`

**Parameters:**  
- `decimal_dollar`: Dollar price in decimal.
- `fraction`: Denominator for fraction.

**Example:**  
```

=DOLLARFR(1.125, 16)

```
Returns: $1.02

---

### DURATION

**Description:**  
Returns annual duration of a security with periodic interest payments.

**Syntax:**  
`=DURATION(settlement, maturity, coupon, yield, frequency, [basis])`

**Parameters:**  
- `settlement`: Date acquired.
- `maturity`: Maturity date.
- `coupon`: Coupon rate.
- `yield`: Annual yield.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=DURATION("01-04-2020", "01-01-2025", 0.08, 0.09, 2)

```
Returns: Duration.

---

### EFFECT

**Description:**  
Returns effective annual interest rate.

**Syntax:**  
`=EFFECT(nominal_rate, npery)`

**Parameters:**  
- `nominal_rate`: Nominal interest rate.
- `npery`: Number of compounding periods/year.

**Example:**  
```

=EFFECT(0.07, 4)

```
Returns: Effective annual rate.

---

### FV

**Description:**  
Returns future value of investment based on periodic payments and constant rate.

**Syntax:**  
`=FV(rate, nper, pmt, [pv], [type])`

**Parameters:**  
- `rate`: Interest rate per period.
- `nper`: Number of periods.
- `pmt`: Payment per period.
- `pv` (optional): Present value.
- `type` (optional): 0=end, 1=beginning.

**Example:**  
```

=FV(0.005, 60, -250, 0, 0)

```
Returns: Future value.

---

### FVSCHEDULE

**Description:**  
Returns future value of initial principal after applying series of interest rates.

**Syntax:**  
`=FVSCHEDULE(principal, schedule)`

**Parameters:**  
- `principal`: Initial investment.
- `schedule`: Range of interest rates.

**Example:**  
```

=FVSCHEDULE(1000, {0.05,0.08,0.09})

```
Returns: FV after 3 years.

---

### INTRATE

**Description:**  
Returns interest rate for fully invested security.

**Syntax:**  
`=INTRATE(settlement, maturity, investment, redemption, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: End date.
- `investment`: Amount invested.
- `redemption`: Value received at maturity.
- `basis` (optional): Day count basis.

**Example:**  
```

=INTRATE("01-04-2020", "01-01-2025", 95000, 100000)

```
Returns: Interest rate.

---

### IPMT

**Description:**  
Returns interest payment for specified period of an investment.

**Syntax:**  
`=IPMT(rate, per, nper, pv, [fv], [type])`

**Parameters:**  
- `rate`: Interest rate per period.
- `per`: Period for calculation.
- `nper`: Total periods.
- `pv`: Present value.
- `fv` (optional): Future value.
- `type` (optional): 0=end, 1=beginning.

**Example:**  
```

=IPMT(0.05/12, 5, 60, 10000)

```
Returns: Interest payment in 5th month.

---

### IRR

**Description:**  
Returns internal rate of return for series of cash flows.

**Syntax:**  
`=IRR(values, [guess])`

**Parameters:**  
- `values`: Array/range of cash flows.
- `guess` (optional): Initial guess.

**Example:**  
```

=IRR({-10000, 2500, 2500, 2500, 2500, 2500})

```
Returns: IRR.

---

### ISPMT

**Description:**  
Returns interest paid during specific period of investment.

**Syntax:**  
`=ISPMT(rate, per, nper, pv)`

**Parameters:**  
- `rate`: Interest rate per period.
- `per`: Specific period.
- `nper`: Number of periods.
- `pv`: Present value.

**Example:**  
```

=ISPMT(0.05/12, 5, 60, 10000)

```
Returns: Interest paid in 5th month.

---

### MDURATION

**Description:**  
Returns modified Macaulay duration for security.

**Syntax:**  
`=MDURATION(settlement, maturity, coupon, yield, frequency, [basis])`

**Parameters:**  
- `settlement`: Acquisition date.
- `maturity`: Maturity date.
- `coupon`: Annual coupon rate.
- `yield`: Annual yield.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=MDURATION("01-04-2020", "01-01-2025", 0.08, 0.09, 2)

```
Returns: Modified duration.

---

### MIRR

**Description:**  
Returns modified internal rate of return.

**Syntax:**  
`=MIRR(values, finance_rate, reinvest_rate)`

**Parameters:**  
- `values`: Array/range of cash flows.
- `finance_rate`: Interest rate paid on money used in cash flows.
- `reinvest_rate`: Interest rate received on reinvestments.

**Example:**  
```

=MIRR({-10000,2500,2500,2500,2500,2500},0.1,0.12)

```
Returns: MIRR.

---

### NOMINAL

**Description:**  
Returns annual nominal interest rate.

**Syntax:**  
`=NOMINAL(effect_rate, npery)`

**Parameters:**  
- `effect_rate`: Effective interest rate.
- `npery`: Compounding periods/year.

**Example:**  
```

=NOMINAL(0.12, 4)

```
Returns: Nominal rate.

---

### NPER

**Description:**  
Returns number of periods for investment.

**Syntax:**  
`=NPER(rate, pmt, pv, [fv], [type])`

**Parameters:**  
- `rate`: Interest rate per period.
- `pmt`: Payment per period.
- `pv`: Present value.
- `fv` (optional): Future value.
- `type` (optional): 0=end, 1=beginning.

**Example:**  
```

=NPER(0.005, -250, 10000)

```
Returns: Number of periods.

---

### NPV

**Description:**  
Returns net present value based on cash flows and discount rate.

**Syntax:**  
`=NPV(rate, value1, [value2], ...)`

**Parameters:**  
- `rate`: Discount rate.
- `value1, value2, ...`: Cash flows.

**Example:**  
```

=NPV(0.1, -10000, 2500, 2500, 2500, 2500, 2500)

```
Returns: Net present value.

---

### ODDFPRICE

**Description:**  
Returns price per $100 face value of security with odd first period.

**Syntax:**  
`=ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, [basis])`

**Parameters:**  
- `settlement`: Settlement date.
- `maturity`: Maturity date.
- `issue`: Issue date.
- `first_coupon`: First coupon date.
- `rate`: Annual coupon rate.
- `yld`: Yield.
- `redemption`: Redemption value per $100.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=ODDFPRICE("01-04-2020", "01-01-2025", "01-01-2020", "01-07-2020", 0.08, 0.09, 100, 2)

```
Returns: Price.

---

### ODDFYIELD

**Description:**  
Returns yield of a security with odd first period.

**Syntax:**  
`=ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, [basis])`

**Parameters:**  
- `settlement`: Settlement date.
- `maturity`: Maturity date.
- `issue`: Issue date.
- `first_coupon`: First coupon date.
- `rate`: Annual coupon rate.
- `pr`: Price per $100 face.
- `redemption`: Redemption value per $100 face.
- `frequency`: Number of payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=ODDFYIELD("01-04-2020", "01-01-2025", "01-01-2020", "01-07-2020", 0.08, 95, 100, 2)

```
Returns: Yield.

---

### ODDLPRICE

**Description:**  
Returns price per $100 face value of a security with odd last period.

**Syntax:**  
`=ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, [basis])`

**Parameters:**  
- `settlement`: Settlement date.
- `maturity`: Maturity date.
- `last_interest`: Last coupon date.
- `rate`: Annual coupon rate.
- `yld`: Yield.
- `redemption`: Redemption value per $100.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=ODDLPRICE("01-04-2020", "01-01-2025", "01-07-2021", 0.08, 0.09, 100, 2)

```
Returns: Price.

---

### ODDLYIELD

**Description:**  
Returns yield of a security with odd last period.

**Syntax:**  
`=ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency, [basis])`

**Parameters:**  
- `settlement`: Settlement date.
- `maturity`: Maturity date.
- `last_interest`: Last coupon date.
- `rate`: Annual coupon rate.
- `pr`: Price per $100 face.
- `redemption`: Redemption value per $100 face.
- `frequency`: Number of payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=ODDLYIELD("01-04-2020", "01-01-2025", "01-07-2021", 0.08, 95, 100, 2)

```
Returns: Yield.

---

### PDURATION

**Description:**  
Returns number of periods required for investment to reach specified value.

**Syntax:**  
`=PDURATION(rate, pv, fv)`

**Parameters:**  
- `rate`: Interest rate per period.
- `pv`: Present value/principal.
- `fv`: Future value.

**Example:**  
```

=PDURATION(0.05, 10000, 20000)

```
Returns: Periods to double investment.

---

### PMT

**Description:**  
Returns periodic payment for an annuity.

**Syntax:**  
`=PMT(rate, nper, pv, [fv], [type])`

**Parameters:**  
- `rate`: Interest rate per period.
- `nper`: Total number of payments.
- `pv`: Present value.
- `fv` (optional): Future value.
- `type` (optional): 0=end, 1=beginning.

**Example:**  
```

=PMT(0.05/12, 60, 10000)

```
Returns: Monthly payment.

---

### PPMT

**Description:**  
Returns payment on the principal for a given period.

**Syntax:**  
`=PPMT(rate, per, nper, pv, [fv], [type])`

**Parameters:**  
- `rate`: Interest rate per period.
- `per`: Period to check.
- `nper`: Total payments.
- `pv`: Present value.
- `fv` (optional): Future value.
- `type` (optional): 0=end, 1=beginning.

**Example:**  
```

=PPMT(0.05/12, 5, 60, 10000)

```
Returns: Principal in period 5.

---

### PRICE

**Description:**  
Returns price per $100 face value of a bond.

**Syntax:**  
`=PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: Maturity date.
- `rate`: Coupon rate.
- `yld`: Yield.
- `redemption`: Redemption value per $100.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=PRICE("01-04-2020", "01-01-2025", 0.08, 0.09, 100, 2)

```
Returns: Bond price.

---

### PRICEDISC

**Description:**  
Returns price per $100 face value of a discounted security.

**Syntax:**  
`=PRICEDISC(settlement, maturity, discount, redemption, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: Maturity date.
- `discount`: Discount rate.
- `redemption`: Redemption value per $100.
- `basis` (optional): Day count basis.

**Example:**  
```

=PRICEDISC("01-04-2020", "01-01-2025", 0.05, 100)

```
Returns: Discounted security price.

---

### PRICEMAT

**Description:**  
Returns price per $100 face value of a security that pays interest at maturity.

**Syntax:**  
`=PRICEMAT(settlement, maturity, issue, rate, yld, redemption, [basis])`

**Parameters:**  
- `settlement`: Settlement date.
- `maturity`: Maturity date.
- `issue`: Issue date.
- `rate`: Annual coupon rate.
- `yld`: Yield.
- `redemption`: Redemption value per $100.
- `basis` (optional): Day count basis.

**Example:**  
```

=PRICEMAT("01-04-2020", "01-01-2025", "01-01-2020", 0.08, 0.09, 100)

```
Returns: Price for security at maturity.

---

### PV

**Description:**  
Returns present value of an investment.

**Syntax:**  
`=PV(rate, nper, pmt, [fv], [type])`

**Parameters:**  
- `rate`: Interest rate per period.
- `nper`: Number of periods.
- `pmt`: Payment per period.
- `fv` (optional): Future value.
- `type` (optional): 0=end, 1=beginning.

**Example:**  
```

=PV(0.005, 60, -250, 0, 0)

```
Returns: Present value.

---

### RATE

**Description:**  
Returns interest rate per period of an annuity.

**Syntax:**  
`=RATE(nper, pmt, pv, [fv], [type], [guess])`

**Parameters:**  
- `nper`: Number of periods.
- `pmt`: Payment per period.
- `pv`: Present value.
- `fv` (optional): Future value.
- `type` (optional): 0=end, 1=beginning.
- `guess` (optional): Initial guess.

**Example:**  
```

=RATE(60, -250, 10000)

```
Returns: Periodic interest rate.

---

### RECEIVED

**Description:**  
Returns amount received at maturity for fully invested security.

**Syntax:**  
`=RECEIVED(settlement, maturity, investment, discount, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: Maturity date.
- `investment`: Amount invested.
- `discount`: Discount rate.
- `basis` (optional): Day count basis.

**Example:**  
```

=RECEIVED("01-04-2020", "01-01-2025", 95000, 0.05)

```
Returns: Amount at maturity.

---

### RRI

**Description:**  
Returns equivalent interest rate for growth of investment.

**Syntax:**  
`=RRI(nper, pv, fv)`

**Parameters:**  
- `nper`: Number of periods.
- `pv`: Present value.
- `fv`: Future value.

**Example:**  
```

=RRI(6, 10000, 20000)

```
Returns: Equivalent rate per period.

---

### SLN

**Description:**  
Returns straight-line depreciation for one period.

**Syntax:**  
`=SLN(cost, salvage, life)`

**Parameters:**  
- `cost`: Asset cost.
- `salvage`: Value after life.
- `life`: Useful life.

**Example:**  
```

=SLN(10000, 1000, 5)

```
Returns: Depreciation per period.

---

### STOCKHISTORY

**Description:**  
Returns historical data for securities.

**Syntax:**  
`=STOCKHISTORY(stock, start_date, end_date, [interval], ...)`

**Parameters:**  
- `stock`: Security identifier (symbol).
- `start_date`: Start date.
- `end_date`: End date.
- `interval` (optional): Data frequency (daily, weekly, monthly).

**Example:**  
```

=STOCKHISTORY("MSFT", "01-01-2020", "12-31-2020")

```
Returns: MSFT history for 2020.

---

### SYD

**Description:**  
Returns sum-of-years' digits depreciation for specified period.

**Syntax:**  
`=SYD(cost, salvage, life, period)`

**Parameters:**  
- `cost`: Initial cost.
- `salvage`: Value after useful life.
- `life`: Number of periods.
- `period`: Period to calculate.

**Example:**  
```

=SYD(10000, 1000, 5, 1)

```
Returns: SYD depreciation for year 1.

---

### TBILLEQ

**Description:**  
Returns bond-equivalent yield for a Treasury bill.

**Syntax:**  
`=TBILLEQ(settlement, maturity, pr)`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: Maturity date.
- `pr`: Price per $100 face value.

**Example:**  
```

=TBILLEQ("01-04-2020", "01-01-2025", 95)

```
Returns: Bond-equivalent yield.

---

### TBILLPRICE

**Description:**  
Returns price per $100 face value for a Treasury bill.

**Syntax:**  
`=TBILLPRICE(settlement, maturity, discount)`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: Maturity.
- `discount`: Discount rate.

**Example:**  
```

=TBILLPRICE("01-04-2020", "01-01-2025", 0.05)

```
Returns: Treasury bill price.

---

### TBILLYIELD

**Description:**  
Returns yield for a Treasury bill.

**Syntax:**  
`=TBILLYIELD(settlement, maturity, pr)`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: End date.
- `pr`: Price per $100 face value.

**Example:**  
```

=TBILLYIELD("01-04-2020", "01-01-2025", 95)

```
Returns: Treasury bill yield.

---

### VDB

**Description:**  
Returns depreciation for asset for specified/partial period using variable declining balance.

**Syntax:**  
`=VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])`

**Parameters:**  
- `cost`: Initial cost.
- `salvage`: Value after useful life.
- `life`: Number of periods.
- `start_period`: Starting period.
- `end_period`: Ending period.
- `factor` (optional): Rate of decline (default 2).
- `no_switch` (optional): TRUE prevents switch to straight-line, FALSE allows.

**Example:**  
```

=VDB(10000, 1000, 5, 1, 2)

```
Returns: Depreciation between periods 1 and 2.

---

### XIRR

**Description:**  
Returns internal rate of return for a schedule of cash flows not necessarily periodic.

**Syntax:**  
`=XIRR(values, dates, [guess])`

**Parameters:**  
- `values`: Array of cash flows.
- `dates`: Dates of cash flows.
- `guess` (optional): Initial guess.

**Example:**  
```

=XIRR({-10000,2500,2500,2500,2500,2500},{"01-01-2020","01-01-2021","01-01-2022","01-01-2023","01-01-2024","01-01-2025"})

```
Returns: IRR.

---

### XNPV

**Description:**  
Returns net present value for a schedule of cash flows not necessarily periodic.

**Syntax:**  
`=XNPV(rate, values, dates)`

**Parameters:**  
- `rate`: Discount rate.
- `values`: Array/range of cash flows.
- `dates`: Dates of cash flows.

**Example:**  
```

=XNPV(0.1, {-10000,2500,2500,2500,2500,2500}, {"01-01-2020","01-01-2021","01-01-2022","01-01-2023","01-01-2024","01-01-2025"})

```
Returns: Net present value.

---

### YIELD

**Description:**  
Returns yield on a security that pays periodic interest.

**Syntax:**  
`=YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: End date.
- `rate`: Coupon rate.
- `pr`: Price per $100 face value.
- `redemption`: Redemption value per $100 face.
- `frequency`: Payments/year.
- `basis` (optional): Day count basis.

**Example:**  
```

=YIELD("01-04-2020", "01-01-2025", 0.08, 95, 100, 2)

```
Returns: Security yield.

---

### YIELDDISC

**Description:**  
Returns annual yield for discounted security (e.g., Treasury bill).

**Syntax:**  
`=YIELDDISC(settlement, maturity, pr, redemption, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: End date.
- `pr`: Price per $100 face value.
- `redemption`: Redemption value per $100.
- `basis` (optional): Day count basis.

**Example:**  
```

=YIELDDISC("01-04-2020", "01-01-2025", 95, 100)

```
Returns: Discounted security yield.

---

### YIELDMAT

**Description:**  
Returns annual yield of security paying interest at maturity.

**Syntax:**  
`=YIELDMAT(settlement, maturity, issue, rate, pr, redemption, [basis])`

**Parameters:**  
- `settlement`: Purchase date.
- `maturity`: End date.
- `issue`: Issue date.
- `rate`: Annual coupon rate.
- `pr`: Price per $100 face value.
- `redemption`: Redemption value per $100 face.
- `basis` (optional): Day count basis.

**Example:**  
```

=YIELDMAT("01-04-2020", "01-01-2025", "01-01-2020", 0.08, 95, 100)

```
Returns: Yield at maturity.

---
```

**Copy and use this file for a complete reference to every major Excel Financial Formula. You can add other categories—DATE \& TIME, MATH \& TRIG, etc.—in exactly this format for a unified Excel guide. If you need the file as raw Markdown text for download or editing, let me know!**
<span style="display:none">[^1][^10][^11][^12][^13][^14][^15][^16][^17][^2][^3][^4][^5][^6][^7][^8][^9]</span>

<div align="center">⁂</div>

[^1]: https://support.microsoft.com/en-us/office/financial-functions-reference-5658d81e-6035-4f24-89c1-fbf124c2b1d8

[^2]: https://www.financialprofessionals.org/training-resources/resources/articles/Details/useful-excel-functions-for-financial-planning-and-analysis

[^3]: https://www.drnishikantjha.com/booksCollection/Financial Analysis with Microsoft Excel ( PDFDrive ).pdf

[^4]: https://courses.lumenlearning.com/wm-computerapplicationsmgrs-2/chapter/financial-functions-and-formulas/

[^5]: https://www.youtube.com/watch?v=XOyWMYQIeIE

[^6]: https://corporatefinanceinstitute.com/resources/excel/excel-for-finance/

[^7]: https://web.acd.ccac.edu/~ndowney/CIT140/Excel/Formulas.pdf

[^8]: https://www.fe.training/free-resources/excel/30-key-excel-functions-for-finance/

[^9]: https://siesce.edu.in/docs/resources/EXCEL Functions_83425.pdf

[^10]: https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188

[^11]: https://openstax.org/books/workplace-software-skills/pages/12-2-financial-functions-in-microsoft-excel

[^12]: https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb

[^13]: https://www.datacamp.com/tutorial/basic-excel-formulas-for-everyone

[^14]: https://www.youtube.com/watch?v=w5cytGoutmg

[^15]: https://www.wallstreetmojo.com/financial-functions-in-excel/

[^16]: https://www.bc.edu/content/dam/files/offices/help-secure/pdf/training/excel-functions.pdf

[^17]: https://www.excel-accountant.com/article/excel-functions-and-formulas-for-finance

