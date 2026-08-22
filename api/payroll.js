import { prisma } from "./_lib/prisma.js";
import { methodGuard, body, errorMessage } from "./_lib/http.js";
const nums=["basic","houseRent","conveyance","food","medical","gross","totalDays","payableDays","present","weekend","leave","absent","paySalary","absentAmount","otRate","otHours","otAmount","actualAmount","advance","arrear","payBeforeTds","tds","payAmount"];
export default async function handler(req,res){
 if(!methodGuard(req,res,["GET","POST","PUT","PATCH"]))return;
 try{
  if(req.method==="GET"){const where={...(req.query?.month?{month:req.query.month}:{}),...(req.query?.employeeId?{employeeId:req.query.employeeId}:{})};return res.status(200).json({records:await prisma.payroll.findMany({where,orderBy:[{month:"desc"},{employeeId:"asc"}]})});}
  const b=body(req);if(!b.employeeId||!b.month)return res.status(400).json({error:"employeeId and month are required"});const d={employeeId:b.employeeId,month:b.month,status:b.status||"Draft"};for(const k of nums)if(b[k]!==undefined)d[k]=Number(b[k])||0;
  return res.status(200).json({record:await prisma.payroll.upsert({where:{employeeId_month:{employeeId:b.employeeId,month:b.month}},create:d,update:d})});
 }catch(e){res.status(500).json({error:errorMessage(e),code:e?.code});}
}
