@echo off
wsl -e bash -lic "conda activate linelist2; cd /mnt/c/apps/linelist; python linelist_generator_gui.py"
