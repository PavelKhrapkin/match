set from=Payment.txt
set to=C:\SFconstr\Payment.csv

echo ********* converting %from% -> %to%

mkdir sed_temp
cd sed_temp
copy ..\%from% temp2.csv
"\Program Files\GnuWin32\bin"\sed -i s/\"//g  temp2.csv

"\Program Files\GnuWin32\bin"\sed -i s/\t/\",\"/g  temp2.csv
"\Program Files\GnuWin32\bin"\sed -i s/$/\"/g temp2.csv
"\Program Files\GnuWin32\bin"\sed -i s/^^/\"/g temp2.csv
cd ..
"\Program Files\GnuWin32\bin"\iconv -c -f CP1251 -t UTF-8 sed_temp\temp2.csv > %to%
del /q sed_temp
rmdir sed_temp
