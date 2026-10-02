#Get image that works for e.g. net481
#FROM mcr.microsoft.com/dotnet/aspnet:10.0
#Get image that works for net8,net9 etc.
FROM mcr.microsoft.com/dotnet/sdk:10.0

#Install mono as dotnet/sdk:10.0 fails otherwise. 
#This is Strange as sdk:9.0 and 8.0 don't seem to require this.
#Look into removing this later. 
#Anyone can test removing it locally by commenting this line out but the whole test project will fail to run properly without it as of 2026-09-28
RUN apt-get update && apt-get install -y mono-complete wget

#This Requires you to have Docker Desktop downloaded and open.
#Create the Visual Studio Debug Shell image.
#Essentially Microsoft has a live link to a shell-script that generates the "vsut_dockerfile" image
RUN wget https://aka.ms/getvsdbgsh && \
    sh getvsdbgsh -v latest  -l /vsdbg