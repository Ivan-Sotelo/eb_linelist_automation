@echo off
wsl -e bash -lic "conda activate linelist2; cd ~/envs/linelist; python linelist_generator_gui.py"