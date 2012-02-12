#!/bin/sh
cd $1
nmake clean
perl Makefile.PL
nmake
tar -cvzf $1-1.tar.gz blib
cd ..

