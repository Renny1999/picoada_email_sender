while read l; do
  grep "$l" files.txt
  
done <skip.txt
