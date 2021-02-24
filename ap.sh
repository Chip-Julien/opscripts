#!/bin/bash
cd /home/robot/Downloads
echo Enter Date "(yyyymmdd)"
read date
echo Creating Folder
mkdir $date
echo Retrieving Data
scp robot@compute01.dc7039.everest.bgrey.io:/home/robot/metrics/$date.tgz /home/robot/Downloads/$date.tgz
mv $date.tgz /home/robot/Downloads/$date
cd /home/robot/Downloads/$date
tar -zxvf $date.tgz
cd .. 
cp make_output.py /home/robot/Downloads/$date
cp template.xlsx /home/robot/Downloads/$date
cd /home/robot/Downloads/$date
python make_output.py -d$date
