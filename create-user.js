import { PrismaClient } from "@prisma/client";
import bcrypt from "bcryptjs";

const prisma = new PrismaClient();

try {
  const email = "admin@example.com";
  const password = "Admin@123";

  const passwordHash = await bcrypt.hash(password, 12);

  const user = await prisma.user.upsert({
    where: {
      email,
    },
    update: {
      passwordHash,
      name: "Administrator",
      role: "Admin",
      status: "Active",
    },
    create: {
      name: "Administrator",
      email,
      passwordHash,
      role: "Admin",
      status: "Active",
    },
  });

  console.log("User ready:");
  console.log("Email:", user.email);
  console.log("Password:", password);
  console.log("Role:", user.role);
} catch (error) {
  console.error("Error creating user:", error);
} finally {
  await prisma.$disconnect();
}