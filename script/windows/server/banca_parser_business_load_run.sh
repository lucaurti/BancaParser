echo "fermo l'esecuzione del container"
sudo docker container stop banca_parser_business_apihost_container	
#sleep 1.5
echo "elimino il container"
sudo docker container rm banca_parser_business_apihost_container
#sleep 1.5
echo "rimuovo l'image"
sudo docker image rm banca_parser_business_apihost_container_image
#sleep 1.5
echo "mi posiziono nella cartella server dove è stato salvato il file"
cd "/home/luclash/upload"
#sleep 1.5
echo "unzip il file"
bunzip2 banca_parser_business_apihost_container_image.tar.bz2
#sleep 1.5
echo "carico l'image in docker"
sudo docker image load --input banca_parser_business_apihost_container_image.tar
#sleep 1.5
echo "eseguo il docker"
sudo docker run -d -p 5600:8080 -v /media/hd/dati500gb/bancaparser/:/BancaParser.ApiHost/Volume_BancaParser_Host/ --name banca_parser_business_apihost_container --network luclash-net -e ASPNETCORE_ENVIRONMENT='server_docker' -e TZ=Europe/Paris --restart unless-stopped -t banca_parser_business_apihost_container_image
#sleep 1.5
echo "effettuo una docker prune"
sudo docker system prune -f
echo "rimuovo i files .tar"
rm *.tar
#sleep 1.5
echo "mi disconetto dal server"
exit