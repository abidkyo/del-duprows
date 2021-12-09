# del_dups_row

## Description

Script to delete duplicates rows by the same values in a column.

- Group the data and apply "first" function.

## Usage

```bash
python3 del_dups_row.py -v -f data.xlsx -g "First Name" -o output.xlsx
```

## Dependencies

[pandas](https://pandas.pydata.org/)
[openpyxl](https://openpyxl.readthedocs.io)

```bash
python3 -m pip install pandas openpyxl
```

## Help Message

```bash
usage: del_dups_row.py [-h] [-v] -f FILE [-s SHEET] [-n NROWS] -g GROUPBY -o OUTPUT

Delete duplicates rows by the same values in a column.

optional arguments:
  -h, --help            show this help message and exit
  -v, --verbose         verbosity
  -s SHEET, --sheet SHEET
                        sheet name or number
  -n NROWS, --nrows NROWS
                        number of rows

required named arguments:
  -f FILE, --file FILE  excel filename
  -g GROUPBY, --groupby GROUPBY
                        groupby column
  -o OUTPUT, --output OUTPUT
                        output filename
```

## Example Output

```
Original Data
   First Name   Last Name  Gender  Age                         Email
0   Charlotte      Walker  Female   21       c.walker@randatmail.com
1      Ashton  Richardson    Male   19   a.richardson@randatmail.com
2       Ellia     Spencer  Female   28      e.spencer@randatmail.com
3       Edwin      Gibson    Male   29       e.gibson@randatmail.com
4     Brianna       Adams  Female   26        b.adams@randatmail.com
5      Amanda      Miller  Female   22       a.miller@randatmail.com
6    Madaline    Sullivan  Female   30     m.sullivan@randatmail.com
7      Victor        Reed    Male   25         v.reed@randatmail.com
8      Gianna       Payne  Female   19        g.payne@randatmail.com
9     Miranda      Martin  Female   20       m.martin@randatmail.com
10      Ellia     Spencer  Female   28  ellia.spencer@randatmail.com
11     Victor        Reed    Male   25    victor.reed@randatmail.com

Data without duplicate rows
  First Name   Last Name  Gender  Age                        Email
0     Amanda      Miller  Female   22      a.miller@randatmail.com
1     Ashton  Richardson    Male   19  a.richardson@randatmail.com
2    Brianna       Adams  Female   26       b.adams@randatmail.com
3  Charlotte      Walker  Female   21      c.walker@randatmail.com
4      Edwin      Gibson    Male   29      e.gibson@randatmail.com
5      Ellia     Spencer  Female   28     e.spencer@randatmail.com
6     Gianna       Payne  Female   19       g.payne@randatmail.com
7   Madaline    Sullivan  Female   30    m.sullivan@randatmail.com
8    Miranda      Martin  Female   20      m.martin@randatmail.com
9     Victor        Reed    Male   25        v.reed@randatmail.com
```

### Note

Tested with Python 3.8.10 and Pandas 1.3.4
