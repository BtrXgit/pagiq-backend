FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY ["DocConverter.csproj", "./"]
RUN dotnet restore "DocConverter.csproj"

COPY . .
RUN dotnet publish "DocConverter.csproj" -c Release -o /app/publish

FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final
WORKDIR /app

# Install Syncfusion fonts, Python, Poppler (for PDF images) & Python Libraries
RUN apt-get update && apt-get install -y libgdiplus fontconfig python3 python3-pip poppler-utils
RUN pip3 install pdf2docx pdfplumber pandas openpyxl pdf2image python-pptx --break-system-packages

COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "DocConverter.dll"]
