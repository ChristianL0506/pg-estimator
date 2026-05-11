import { build as esbuild } from "esbuild";
import { build as viteBuild } from "vite";
import { rm, readFile, copyFile, mkdir } from "fs/promises";
import { existsSync } from "fs";

// server deps to bundle to reduce openat(2) syscalls
// which helps cold start times
const allowlist = [
  "@google/generative-ai",
  "axios",
  "cors",
  "date-fns",
  "drizzle-orm",
  "drizzle-zod",
  "express",
  "express-rate-limit",
  "express-session",
  "jsonwebtoken",
  "memorystore",
  "multer",
  "nanoid",
  "nodemailer",
  "openai",
  "passport",
  "passport-local",
  "stripe",
  "uuid",
  "ws",
  "xlsx",
  "zod",
  "zod-validation-error",
];

async function buildAll() {
  await rm("dist", { recursive: true, force: true });

  console.log("building client...");
  await viteBuild();

  console.log("building server...");
  const pkg = JSON.parse(await readFile("package.json", "utf-8"));
  const allDeps = [
    ...Object.keys(pkg.dependencies || {}),
    ...Object.keys(pkg.devDependencies || {}),
  ];
  const externals = allDeps.filter((dep) => !allowlist.includes(dep));

  await esbuild({
    entryPoints: ["server/index.ts"],
    platform: "node",
    bundle: true,
    format: "cjs",
    outfile: "dist/index.cjs",
    define: {
      "process.env.NODE_ENV": '"production"',
    },
    minify: true,
    external: externals,
    logLevel: "info",
  });

  // Copy runtime data files that the server reads via fs.readFileSync.
  // These must live next to dist/index.cjs because the server uses
  // path.join(__dirname, "<file>") to find them — and __dirname resolves
  // to dist/ at runtime. Without this copy step, calls to getEstimatorData()
  // fall through to a CWD-based fallback that breaks if Node is launched
  // from anywhere except the repo root.
  const runtimeFiles = [
    { from: "server/estimator-data.json", to: "dist/estimator-data.json" },
  ];
  for (const f of runtimeFiles) {
    if (existsSync(f.from)) {
      await copyFile(f.from, f.to);
      console.log(`copied ${f.from} → ${f.to}`);
    } else {
      console.warn(`runtime file not found: ${f.from} — skipping`);
    }
  }
}

buildAll().catch((err) => {
  console.error(err);
  process.exit(1);
});
