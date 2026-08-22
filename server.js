import express from "express";
import { PrismaClient } from "@prisma/client";
import bcrypt from "bcryptjs";

const app = express();
const prisma = new PrismaClient();

app.use(express.json());

app.post("/api/auth", async (req, res) => {
  try {
    const { email, password } = req.body || {};

    if (!email || !password) {
      return res.status(400).json({
        success: false,
        message: "Email and password are required.",
      });
    }

    const loginValue = String(email).trim().toLowerCase();

    const user = await prisma.user.findUnique({
      where: {
        email: loginValue,
      },
    });

    if (!user) {
      return res.status(401).json({
        success: false,
        message: "Invalid username/email or password.",
      });
    }

    const passwordMatch = await bcrypt.compare(
      String(password),
      user.passwordHash
    );

    if (!passwordMatch) {
      return res.status(401).json({
        success: false,
        message: "Invalid username/email or password.",
      });
    }

    if (user.status !== "Active") {
      return res.status(403).json({
        success: false,
        message: "This account is inactive.",
      });
    }

    return res.json({
      success: true,
      message: "Login successful.",
      user: {
        id: user.id,
        name: user.name,
        email: user.email,
        role: user.role,
        status: user.status,
      },
    });

  } catch (error) {
    console.error("AUTH ERROR:", error);

    return res.status(500).json({
      success: false,
      message: "Authentication server error.",
    });
  }
});

app.get("/api/health", (req, res) => {
  res.json({
    success: true,
    message: "HRM API is running",
  });
});

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`HRM API running on http://localhost:${PORT}`);
});