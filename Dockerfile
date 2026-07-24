# ============================================
# Dockerfile — Asset Management System
# Multi-stage build for .NET 10
# ============================================

# Stage 1: Build
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src

# Copy project file and restore dependencies
COPY ["Asset_Management.csproj", "."]
RUN dotnet restore "Asset_Management.csproj"

# Copy all source code
COPY . .

# Build and publish in release mode
RUN dotnet publish "Asset_Management.csproj" -c Release -o /app/publish

# Stage 2: Runtime
FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime
WORKDIR /app

# Copy published artifacts from build stage
COPY --from=build /app/publish .

# Expose ports
EXPOSE 8080

# Set environment
ENV ASPNETCORE_ENVIRONMENT=Production
ENV ASPNETCORE_URLS=http://+:8080

# Run the application
ENTRYPOINT ["dotnet", "Asset_Management.dll"]
