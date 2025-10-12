@echo off
echo "start"
echo "mi posiziono nella cartella padre di tutti i progetti"
cd "C:\luclash\BancaParser\console\"

echo "rimuovo ricorsivamente gli oggetti"
rm -rf ./obj
rm -rf ./bin

echo "mi posiziono nella cartella padre di tutti i progetti"
cd "C:\luclash\BancaParser\console\BancaParser.Console\"

echo "rimuovo ricorsivamente gli oggetti"
rm -rf ./obj
rm -rf ./bin

echo "mi posiziono nella cartella dove creare la cartella per i tar"
cd "C:\temp\"

echo "rimuovo la cartella tar"
rm -rf ./tar

echo "creo la cartella dove salvare i file"
mkdir "tar"

echo "creo la cartella dove salvare i file di publish"
cd "C:\temp\tar"
mkdir "banca-parser"

echo "mi posiziono nella cartella dove è presente il progetto da pubblicare"
cd "C:\luclash\BancaParser\console\BancaParser.Console\"

echo "effettuo l'operazione di publish"
dotnet publish -c Release --self-contained true -r linux-x64 "BancaParser.Console.csproj" -o "C:\temp\tar\banca-parser"

echo "mi posiziono nella cartella dove creare la cartella per i tar"
cd "C:\temp\tar"

echo "creo un file tar"
tar -cvf banca-parser.tar banca-parser

echo "zippo il file creato precedentemente"
bzip2 banca-parser.tar

echo "effettuo l'upload del file"
set RETRIES=5
set COUNT=0

:retry
set /A COUNT+=1
echo Tentativo %COUNT%...
scp C:\temp\tar\banca-parser.tar.bz2 luclash@luclash-server:/home/luclash/upload/banca-parser.tar.bz2
if %ERRORLEVEL% EQU 0 goto :done
if %COUNT% GEQ %RETRIES% goto :fail
echo "SCP Fallito". Riprovo tra 5 secondi...
timeout /t 5 > nul
goto :retry

:fail
echo Operazione fallita dopo %RETRIES% tentativi.
exit /b 1

:done
echo "eseguo lo script per deployare il container"
ssh -tt luclash@luclash-server "sudo bash -c 'dos2unix /home/luclash/app/banca-parser/script/*.sh && bash /home/luclash/app/banca-parser/script/publish_banca-parser_console_server.sh'"

echo "mi posiziono nella cartella dove è presente la cartella per i tar"
cd "C:\temp\"

echo "rimuovo la cartella tar"
rm -rf ./tar

echo "end"