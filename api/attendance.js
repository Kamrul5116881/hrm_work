import { prisma } from "./_lib/prisma.js";
import { methodGuard, body, errorMessage } from "./_lib/http.js";
const fields=["employeeId","month","present","weekend","leave","absent","otHours","advance","arrear","tds","basic","houseRent","medical","conveyance","food","gross","totalDays","payableDays"];
const nums=new Set(fields.slice(2));
function data(x){const d={};for(const k of fields)if(x[k]!==undefined)d[k]=nums.has(k)?Number(x[k])||0:x[k];return d;}
export default async function handler(req,res){
 if(!methodGuard(req,res,["GET","POST","PUT","PATCH","DELETE"]))return;
 try{
  if(req.method==="GET"){const where={...(req.query?.month?{month:req.query.month}:{}),...(req.query?.employeeId?{employeeId:req.query.employeeId}:{})};return res.status(200).json({records:await prisma.attendance.findMany({where,orderBy:[{month:"desc"},{employeeId:"asc"}]})});}
  const b=body(req);if(!b.employeeId||!b.month)return res.status(400).json({error:"employeeId and month are required"});
  const where={employeeId_month:{employeeId:b.employeeId,month:b.month}};
  if(req.method==="DELETE"){await prisma.attendance.delete({where});return res.status(200).json({ok:true});}
  return res.status(200).json({record:await prisma.attendance.upsert({where,create:data(b),update:data(b)})});
 }catch(e){res.status(e?.code==="P2025"?404:500).json({error:errorMessage(e),code:e?.code});}
}
