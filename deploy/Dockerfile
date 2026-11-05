# Optional: GEIR in a container. Build from the unzipped bundle (release/geir-<version>/):
#   docker build -f deploy/Dockerfile -t geir .
#   docker run -d --name geir -p 127.0.0.1:4300:4300 --env-file .env geir
# The Playwright image already has Chromium and its Linux libraries.
FROM mcr.microsoft.com/playwright:v1.61.1-noble

ENV NODE_ENV=production \
    GEIR_HOST=0.0.0.0 \
    GEIR_PORT=4300 \
    GEIR_BROWSER=chromium \
    GEIR_ALLOW_LOCAL_PATHS=false \
    GEIR_TRUST_PROXY=true

# Compiling branches (admin option "Rama"): git, JDK 8 and Maven next to Node.
USER root
RUN apt-get update \
    && apt-get install -y --no-install-recommends git openjdk-8-jdk-headless maven \
    && rm -rf /var/lib/apt/lists/*
ENV GEIR_JAVA_HOME=/usr/lib/jvm/java-8-openjdk-amd64 \
    GEIR_REPOS_DIR=/opt/geir/repos

WORKDIR /opt/geir
COPY . .
RUN mkdir -p /opt/geir/repos && chown -R pwuser /opt/geir/repos
USER pwuser
EXPOSE 4300
CMD ["node", "src/server/server.js"]
