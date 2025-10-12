echo "mi posiziono nella cartella server dove è stato salvato il file"
cd "/home/luclash/app/banca-parser"
echo "Rimuovo la cartella console"
rm -rf ./console
echo "Creo la cartella console"
mkdir console
echo "Copio il file"
cp "/home/luclash/upload/banca-parser.tar.bz2" "/home/luclash/app/banca-parser/console/banca-parser.tar.bz2"
echo "mi posiziono nella cartella server dove è stato salvato il file"
cd "/home/luclash/app/banca-parser/console"
echo "decomprimo il file"
bunzip2 banca-parser.tar.bz2
echo "scompongo il file tar"
tar -xvf banca-parser.tar
echo "rimuovo i files .tar"
rm *.tar
echo "rimuovo il file bz2"
rm "/home/luclash/upload/banca-parser.tar.bz2"
echo "mi disconetto dal server"
exit