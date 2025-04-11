#!/bin/sh

# fc-list
# echo
mkdir -p /home/runner/.fonts
cp -r ./fonts/* /home/runner/.fonts
echo
fc-cache -fv
# echo
# fc-list

#rm -rf fonts/