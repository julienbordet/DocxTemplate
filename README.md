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
  Courriers augmentation 2021-22.docx    -> the Word model 
```

# Installation

The examples below are for year 2023. Adapt depending on the year

## install.sh

Simply launch

```
$ chmod u+x install.sh
$ ./install.sh 2023
```


## Manual way (don't do it) 
### Directory structure

```shell
$ mkdir -p target/2023
```
### Script config

Adapt basic script variables in `config.py`, mainly changing the `year` variable to 2023.

```python
# 
year = "2023"
```

### Adapt docx template

In `template/`, adapt the model file, while being careful to change its name to reflect the year (for instance, in 2023 it should end by `2022-2023.docx`).

### Edit raises and bonus
Copy `template/Augmentation Primes 2022-2023.xlsx` into `target/2023` (while changing the year in both in name of the Excel file and the directory, to be
coherent with the `year` variable above.

Edit the new `template/Augmentation Primes 2022-2023.xlsx` file while respecting the column titles.

**Warning** : `raise` values should be *integers*.
