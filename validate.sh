TESTFILE=`pwd`/$1
echo "Testing file $TESTFILE"

OUTPUT=`docker run --rm -v $TESTFILE:/TestFile.xlsx vindvaki/xlsx-validator /usr/local/bin/xlsx-validator /TestFile.xlsx`

if [ -z "$OUTPUT" ]
  then
    echo "===> Package passes validation"
  else
    echo "===> Package has errors"
    echo "$OUTPUT"
  fi
  