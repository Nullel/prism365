FROM node:24-bookworm-slim AS build
WORKDIR /app
COPY . .
# Tenant data (a PGlite directory, a backup) or a key in the build context would be copied into
# the image. .dockerignore keeps them out; this stops the build if one gets through anyway.
RUN found="$(find . -path '*/node_modules' -prune -o -type f \( -iname PG_VERSION -o -iname prism365.lock -o -iname '*.pem' -o -iname '*.key' -o -iname '*.pfx' -o -iname '*.p12' -o -iname '*.dump' -o -iname '*.dump.*' -o \( \( -iname '*.sql' -o -iname '*.sql.*' \) ! \( -path ./packages/storage/migrations/001-state.sql -o -path ./packages/storage/migrations/002-modules.sql -o -path ./packages/storage/migrations/003-attention.sql -o -path ./packages/storage/migrations/004-audit-runs.sql -o -path ./packages/storage/migrations/006-handoffs.sql \) \) -o -iname '*.tar' -o -iname '*.tar.*' -o -iname '*.tgz' -o -iname '*.zip' -o -iname '*.7z' -o -iname '*.gz' -o -iname '*.zst' -o -iname '*.rar' -o -iname '*.xz' -o -iname '*.bz2' -o -iname '*.backup' -o -iname '*.bak' -o -iname '*.pgdump' -o -iname '*.custom' -o -iname '*.dat' -o -iname '*.lz4' -o -iname 'id_rsa*' -o -iname 'id_dsa*' -o -iname 'id_ecdsa*' -o -iname 'id_ed25519*' -o -iname '.env' -o -iname '.envrc' -o -iname '.env-*' -o -iname '*.env' -o -iname '.npmrc' -o -iname '*.csv' -o \( -iname '.env.*' ! -name '.env.example' \) \) -print)"; \
    if [ -n "$found" ]; then echo "Refusing to build: tenant data or keys in the build context:" >&2; echo "$found" >&2; exit 1; fi
RUN npm ci --ignore-scripts && npm run build

FROM node:24-bookworm-slim AS runtime
WORKDIR /app
COPY --from=build /app/package.json /app/package-lock.json ./
COPY --from=build /app/apps ./apps
COPY --from=build /app/packages ./packages
RUN npm ci --omit=dev --ignore-scripts && npm cache clean --force
ENV NODE_ENV=production HOST=0.0.0.0 PORT=8787 DATA_MODE=demo PUBLIC_ORIGIN=http://localhost:8080 WEB_DIST_PATH=/app/apps/web/dist
USER node
EXPOSE 8787
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s CMD node -e "fetch('http://127.0.0.1:8787/health').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"
CMD ["node", "apps/api/src/server.ts"]
