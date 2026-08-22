import { prisma } from "./_lib/prisma.js";
import { methodGuard, errorMessage } from "./_lib/http.js";
export default async function handler(req,res){
  if(!methodGuard(req,res,["GET"])) return;
  try { await prisma.$queryRaw`SELECT 1`; res.status(200).json({ok:true,database:"connected",service:"hrm-api"}); }
  catch(e){ res.status(500).json({ok:false,database:"error",error:errorMessage(e)}); }
}
