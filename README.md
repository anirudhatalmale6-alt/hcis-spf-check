# spfcheck

A check for the HCIS monthly **SPF DED** file, to run before it goes to the
Finance Department.

It reads the file and nothing else. It never writes to it, never changes a
figure, and never sends it anywhere. It prints what it found and stops.

No payroll data lives in this repository — only the script.

## Running it

```
pip install openpyxl
python3 spfcheck.py "SPF DED SPF AUGUST 2026.xlsx"
```

It exits with 0 when there is nothing to flag and 1 when there is, so it can
be wired into something later if that turns out to be useful.

## What it checks

| | |
|---|---|
| Columns | the twelve expected headings, in the expected order |
| Duplicates | the same NIN more than once, and what the repeats are worth |
| NIN shape | a NIN not written the same way as all the others |
| One person, two NINs | the same name appearing under two different NINs |
| Arithmetic | basic + allowances − deductions = net, on every row |
| Pension | 5% of allowable pay, employee and employer equal |
| Days withheld | whether a reduction in pensionable pay is a whole number of days at basic ÷ days in the month |
| Formats | a money column wearing a date or a time format |
| Empty rows | thousands of blank rows below the data that make the sheet look longer than it is |
| One period | every row carrying the same year and month |

## Why the NIN checks matter here

Finance do not register a carer before paying them. They take the NIN and the
name from this file and process it. There is no master list on their side to
check against.

That makes this file the **only** place a wrong NIN can be caught. A malformed
one sends money to an address nobody owns; the same person under two NINs is
paid twice; and neither is visible downstream, because downstream has nothing
to compare against.

## What it cannot check

It cannot tell you **why** anything was deducted, because the file does not
say. `Deductions` is a single figure that bundles the 5% pension contribution
together with absence and with anything else — a salary advance being repaid,
for example. Ten days absent and a flat 2,000 repayment reach that column
looking identical.

So a row this script flags as "not a whole number of days" is not necessarily
wrong. It means nobody reading the file can tell what the number is made of,
which is worth knowing in its own right.

## Notes on the August 2026 file

Verified against the real file, 3,429 people. Everything below is what the
file itself says, not an assumption:

- Every row adds up. The arithmetic is sound throughout.
- The pension contribution is exactly 5% of allowable pay, employee and
  employer alike, on every row.
- Basic pay takes one of two values only — one for Full Day, one for Half Day.
  There is no day count and no daily rate anywhere in the file.
- 203 people of 3,429 (5.9%) had pay withheld. Of those, 186 land exactly on a
  whole number of days at basic ÷ 31, which is how August's daily rate was
  worked out. The remaining 17 do not, and some of those are round figures
  that look more like flat amounts than day counts.

The NIN-shape and one-person-two-NINs checks were added after that run and
have **not** been tried against the real file - only against files built to
contain each fault and files built to contain none. Run it on the August file
to see what they say.

The alarms have been tested in both directions — on a file built to contain
each fault, and on a file built to contain none — because a check that only
ever says "fine" is indistinguishable from a check that is broken. Doing that
found a fault in this script that the real file never triggered.
