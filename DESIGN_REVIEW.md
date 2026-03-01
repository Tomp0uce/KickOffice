# Design Review v3 — Status

**Date**: 2026-03-01 | **Scope**: Full codebase audit (security, logic, quality, infra)
**Progress**: 🟢 131 implemented · 🔴 0 remaining · 🟡 4 deferred · **135 total**

---

## 🔴 Outstanding Items

*All outstanding items have been implemented.*

---

## 🟡 Deferred Items

**IC2** · `frontend/Dockerfile`, `backend/Dockerfile`
Containers run as root (Node.js default). Adding `USER node` is the best-practice fix but retained intentionally for this deployment.

**IH2** · `frontend/Dockerfile`
Private IP address baked into the default `VITE_BACKEND_URL` build arg — leaks in the compiled JS.
Retained as-is; users must override at build time.

**IH3** · `.env.example`
External DuckDNS domain used as default example value. Retained; users replace with their own values.

**UM10** · `frontend/src/utils/powerpointTools.ts`
Character-by-character HTML reconstruction for PowerPoint — high complexity, low ROI. Deferred pending PowerPoint feature priority.

---

## 🟢 Implemented (131 items)

### Batch 2 (Architecture, Infra, Security, etc)
🟢 BH6 · UC2 · XH1 · EM1 · EM2 · UM1 · UM2 · CH5 · PL2 · PL3 · CM3 · XM1 · BM6 · BM10 · UL3 · CL2 · PM2 · PM5 · PM10 · TM1 · TM2 · UL4 · AL1 · TL2 · IM8 · IL3

### Backend
🟢 BC1–BC4 · BH1–BH5 · BH7 · BM1–BM5 · BM7–BM9 · BL1–BL4

### Frontend Utils
🟢 UC1 · UC3 · UH1–UH7 · UM3–UM9 · UL1 · UL2

### Composables
🟢 CC1 · CC2 · CH1–CH4 · CH6 · CH7 · CM1 · CM2 · CM4–CM11 · CL1 · CL3–CL5

### Infrastructure
🟢 IC1 · IC3 · IH1 · IH4 · IH5 · IM1–IM7 · IL1 · IL2 · IL4–IL7

### Pages / Components / API
🟢 PC1 · PH1–PH5 · AH1 · AH2 · PM1 · PM3 · PM4 · PM6–PM9 · PM11 · AM1–AM4 · PL1 · PL4 · PL5 · TL1 · XM2
