@echo off
echo "start"
echo "mi posiziono nella cartella padre di tutti i progetti"
cd "/Users/lucaurti/Develop/BancaParser/server/BancaParser.ApiHost/"

echo "rimuovo ricorsivamente gli oggetti"
rm -rf ./obj
rm -rf ./bin

echo "mi posiziono nella cartella padre di tutti i progetti"
cd "/Users/lucaurti/Develop/BancaParser/server/BancaParser.Core/"

echo "rimuovo ricorsivamente gli oggetti"
rm -rf ./obj
rm -rf ./bin

echo "mi posiziono nella cartella dove è presente il progetto da pubblicare"
cd "/Users/lucaurti/Develop/BancaParser/server/BancaParser.ApiHost/"

echo "effettuo l'operazione di publish"
dotnet publish -c Release --self-contained true -r linux-x64 "BancaParser.ApiHost.csproj"

echo "copio il dockerfile"
cp "./Dockerfile" "bin/Release/net9.0/linux-x64/"

echo "rimuovo l'image docker già presente"
docker image rm banca_parser_business_apihost_container_image

echo "effettuo la build dell'image"
docker build --platform linux/x86_64 -t banca_parser_business_apihost_container_image .

echo "mi posiziono nella cartella dove creare la cartella per i tar"
cd "/Users/lucaurti/Develop/fileBrowser"

echo "rimuovo la cartella tar"
rm -rf ./tar

echo "creo la cartella dove salvare il file tar"
mkdir "tar"

echo "mi posiziono nella cartella dove salvare il file tar"
cd "./tar"

echo "salvo nella cartella l'image creata in formato tar"
docker image save banca_parser_business_apihost_container_image > banca_parser_business_apihost_container_image.tar 

echo "zippo il file creato precedentemente"
bzip2 banca_parser_business_apihost_container_image.tar

echo "effettuo l'upload del file"
set RETRIES=5
set COUNT=0
while true; do
  COUNT=$((COUNT + 1))
  echo "Tentativo $COUNT..."

  scp /Users/lucaurti/Develop/fileBrowser/tar/banca_parser_business_apihost_container_image.tar.bz2 luclash@luclash-server:/home/luclash/upload/banca_parser_business_apihost_container_image.tar.bz2

  if [ $? -eq 0 ]; then
    break
  fi

  if [ "$COUNT" -ge "$RETRIES" ]; then
    echo "Operazione fallita dopo $RETRIES tentativi."
    exit 1
  fi

  echo "SCP fallito. Riprovo tra 5 secondi..."
  sleep 5
done

echo "eseguo lo script per deployare il container"
ssh -tt luclash@luclash-server "sudo bash -c 'dos2unix /home/luclash/docker/banca-parser-app/script/*.sh && bash /home/luclash/docker/banca-parser-app/script/banca_parser_business_load_run.sh'"

echo "mi posiziono nella cartella dove è presente la cartella per i tar"
cd "/Users/lucaurti/Develop/fileBrowser/"

echo "rimuovo la cartella tar"
rm -rf ./tar

echo "end"