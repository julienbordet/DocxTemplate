#!/usr/bin/env bash

template_core_letter="Courriers-augmentation"
template_core_data="Augmentation-Primes"

function usage() {
  echo "$0 year"
  echo "   year : four digit current year for payraise and bonus"
  echo "          for instance for 2022-2023 payraise, enter 2023"
}

if [ "$1" == "" ]; then
  usage
  exit
fi

year=$1
previous_year=$((year - 1))


if [ ! -d ./venv ]; then
  echo "##### Setting up virtualenv #####"
  python3 -m pip install virtualenv
  virtualenv -p python3 venv
fi

. ./venv/bin/activate

echo "##### Installing requirements #####"
pip3 install -r requirements.txt 2>/dev/null

echo "##### Creating directory structures and template files"

echo " --> Creating target dir"
mkdir -p "target/$year"

echo " --> Updating script"
perl -pi -e "s/2022/$year/g" app/config.py

echo " --> Instancing template"
target_docx="templates/$template_core_letter-$previous_year-$year.docx"
cp "templates/$template_core_letter-2021-2022.docx" "$target_docx"
echo "     Created into templates/$template_core_letter $previous_year-$year"

echo " --> Instancing raise and bonus data"
target_xlsx="target/$year/$template_core_data-$previous_year-$year.xlsx"
cp "templates/$template_core_data-2021-2022.xlsx" "$target_xlsx"

echo ""
echo "Now edit $target_xlsx to configure raises and bonuses "
echo "and $target_docx to setup letter model"
echo ""
echo "then go back to base directory and launch"
echo "  python3 app/payraise.py"
