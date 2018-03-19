#!/bin/sh
basepath=$(cd `dirname $0`; pwd)
chmod u+x ./export.sh
node index.js --export