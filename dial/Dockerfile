# hq-dial — renders the All Orders banner dial as a PNG.
#
# ⚠ DEBIAN, NOT ALPINE. @napi-rs/canvas ships prebuilt native binaries; the glibc build is
#   the well-trodden one. On musl a missing prebuild does not fail loudly — it falls back to
#   compiling from source inside the image, which turns a 20-second build into a toolchain
#   problem. Not worth the ~40 MB.
#
# ⚠ NO SYSTEM FONTS ARE INSTALLED, deliberately. The VPS has no Oswald (fc-list -> 0 faces,
#   checked 2026-08-31) and nobody will remember to add it after the next rebuild. The six
#   static weights live in src/fonts and are registered by path at boot — so the image is
#   self-contained and the drawing cannot silently fall back to a system sans.
FROM node:22-slim

ENV NODE_ENV=production \
    PORT=8099

WORKDIR /app

# Dependencies first, so a change to the drawing does not re-resolve the tree.
# ⚠ `npm ci` and not `npm install` — it installs exactly what the lockfile pins and fails
#   if the lockfile and package.json disagree, which is the whole point of shipping one.
COPY package.json package-lock.json ./
RUN npm ci --omit=dev && npm cache clean --force

COPY src ./src

# ⚠ The node image already ships an unprivileged `node` user. Nothing here writes to disk,
#   so the whole app is owned by root and merely readable — a compromised renderer cannot
#   modify its own code.
USER node

EXPOSE 8099

# ⚠ NO curl IN slim, and none is added — node can ask the question itself. The check reads
#   the BODY, not the status: this box answers 200 with the Floor Board's HTML for anything
#   unrouted, so a status code proves nothing here.
HEALTHCHECK --interval=30s --timeout=4s --start-period=5s --retries=3 \
  CMD node -e "require('http').get({host:'127.0.0.1',port:process.env.PORT||8099,path:'/healthz',timeout:3000},r=>{let b='';r.on('data',c=>b+=c);r.on('end',()=>process.exit(/\"service\":\"hq-dial\"/.test(b)?0:1))}).on('error',()=>process.exit(1))"

CMD ["node", "src/server.js"]
