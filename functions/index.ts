// functions/src/index.ts
import * as functions from "firebase-functions";
import * as admin from "firebase-admin";
import express from "express";
import cors from "cors";
import { v4 as uuidv4 } from "uuid";

admin.initializeApp();
const db = admin.firestore();

const app = express();
// Allow cross-origin requests from your domains; adjust in production
app.use(cors({ origin: true }));
app.use(express.json());

const DEFAULT_TRIAL_USES = 5;

app.post("/requestTrial", async (req, res) => {
  try {
    const { email } = req.body || {};
    const token = uuidv4();
    const doc = {
      token,
      usesLeft: DEFAULT_TRIAL_USES,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      email: email || null,
      lastUsedAt: null,
    };
    await db.collection("trials").doc(token).set(doc);
    res.json({ token, usesLeft: DEFAULT_TRIAL_USES });
  } catch (err: any) {
    console.error(err);
    res.status(500).json({ error: err.message || "requestTrial failed" });
  }
});

app.post("/validateTrial", async (req, res) => {
  try {
    const { token, consume } = req.body || {};
    if (!token) return res.status(400).json({ valid: false, message: "token required" });

    const ref = db.collection("trials").doc(token);
    const snap = await ref.get();
    if (!snap.exists) return res.json({ valid: false });

    const data = snap.data() as any;
    if (!data) return res.json({ valid: false });

    // if not consuming, just return current uses
    if (!consume) {
      return res.json({ valid: data.usesLeft > 0, usesLeft: data.usesLeft || 0 });
    }

    // Atomically decrement if usesLeft > 0
    const txResult = await admin.firestore().runTransaction(async (tx) => {
      const s = await tx.get(ref);
      const current = s.data()?.usesLeft ?? 0;
      if (current <= 0) return { ok: false, usesLeft: 0 };
      tx.update(ref, {
        usesLeft: admin.firestore.FieldValue.increment(-1),
        lastUsedAt: admin.firestore.FieldValue.serverTimestamp(),
      });
      return { ok: true, usesLeft: current - 1 };
    });

    if (!txResult.ok) return res.json({ valid: false, usesLeft: 0 });
    return res.json({ valid: true, usesLeft: txResult.usesLeft });
  } catch (err: any) {
    console.error(err);
    res.status(500).json({ error: err.message || "validateTrial failed" });
  }
});

// export as a single HTTPS function
export const api = functions.https.onRequest(app);
