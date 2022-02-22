#!/usr/bin/env bash

template_core_letter="STIMIO - Courriers augmentation"
template_core_data="Augmentation Primes"

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
mkdir -p target/$year

echo " --> Updating script"
perl -pi -e "s/2022/$year/g" app/payraise.py

echo " --> Instancing template"
cp "templates/$template_core_letter 2021-2022.docx" "templates/$template_core_letter $previous_year-$year"
echo "     Created into templates/$template_core_letter $previous_year-$year"

echo " --> Instancing raise and bonus data"
cp "templates/$template_core_data 2021-2022.xlsx" "target/$year/$template_core_data $previous_year-$year"
echo "     Ready to edit in target/$year/$template_core_data $previous_year-$year"
