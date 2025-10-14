FROM mcr.microsoft.com/dotnet/aspnet:9.0 AS base

VOLUME ["/BancaParser.ApiHost/Volume_BancaParser_ApiHost"]
COPY bin/Release/net9.0/linux-x64/ BancaParser.ApiHost/
RUN chmod 755 /BancaParser.ApiHost/BancaParser.ApiHost

ENTRYPOINT ["./BancaParser.ApiHost/BancaParser.ApiHost"]
