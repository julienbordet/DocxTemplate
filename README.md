# PayRaise
Simple Python script to generate letter for PayRaise from an Excel file.

# Expected directory structure

```
app/
  payraise.py                                     -> the main script

target/
  2022/                                           -> one directory for each year
  2022/Augmentation Primes 2021-2022.xlsx         -> the file containing the raises and bonus

template/
  STIMIO - Courriers augmentation 2021-22.docx    -> the Word model 
```

# Installation

## Directory structure

```shell
$ mkdir -p target/2022
```
## Script config

Adapt basic script variables in `payraise.py`, mainly changing the `year` variable.

```python
[...]
docx_model_name = "templates/STIMIO - Courriers augmentation 2021-22.docx"
year = "2022"
target_dir = f"target/{year}"
raise_xlsx_name = f"Augmentation Primes {int(year)-1}-{year}.xlsx"
[...]
```

## Adapt docx template

In `template/`, adapt the model file, while being careful to change its name to reflect the year (for instance, in 2023 it should end by `2022-2023.docx`).

## Edit raises and bonus
Copy `template/Augmentation Primes 2021-2022.xlsx` into `target/2022` (while changing the year in both in name of the Excel file and the directory, to be
coherent with the `year` variable above.

Edit the new `template/Augmentation Primes 2021-2022.xlsx` file while respecting the column. 

Warning : raise value should be *integers*.
