import { prisma } from "./_lib/prisma.js";
import { methodGuard, body, errorMessage } from "./_lib/http.js";
const ef=["employeeId","name","joiningDate","jobTitle","section","basic","houseRent","conveyance","food","medical","gross","casualLeaveAlloc","medicalLeaveAlloc","motherName","fatherName","dob","nid","maritalStatus","status"];
const af=["employeeId","month","present","weekend","leave","absent","otHours","advance","arrear","tds","basic","medical","conveyance","food","gross","totalDays","payableDays"];
const nums=new Set(["basic","houseRent","conveyance","food","medical","gross","casualLeaveAlloc","medicalLeaveAlloc","present","weekend","leave","absent","otHours","advance","arrear","tds","totalDays","payableDays"]);
function pick(x,fs){const d={};for(const k of fs)if(x[k]!==undefined)d[k]=nums.has(k)?Number(x[k])||0:x[k];return d;}
export default async function handler(req,res){
 if(!methodGuard(req,res,["GET","POST"]))return;
 try{
  if(req.method==="GET"){
   const [employees,attendanceRecords,approvals,rules]=await Promise.all([prisma.employee.findMany({orderBy:{employeeId:"asc"}}),prisma.attendance.findMany({orderBy:[{month:"desc"},{employeeId:"asc"}]}),prisma.payrollApproval.findMany(),prisma.hrRule.findUnique({where:{id:1}})]);
   return res.status(200).json({state:{employees:employees.map(({createdAt,updatedAt,...e})=>e),attendanceRecords:attendanceRecords.map(({createdAt,updatedAt,...a})=>a),payrollApprovals:Object.fromEntries(approvals.map(a=>[a.month,{approved:a.approved,approvedAt:a.approvedAt?.toISOString()||null}])),rules:rules||{payDaysDivisor:31,absentDaysDivisor:30,otDivisor:104,basicDivisor:1.5}}});
  }
  const b=body(req),s=b.state||b;if(!s||!Array.isArray(s.employees)||!Array.isArray(s.attendanceRecords))return res.status(400).json({error:"Expected state.employees and state.attendanceRecords"});
  await prisma.$transaction(async tx=>{
   const ids=new Set(s.employees.map(e=>e.employeeId).filter(Boolean));
   for(const raw of s.employees){if(!raw.employeeId||!raw.name)continue;const d=pick(raw,ef),old=await tx.employee.findUnique({where:{employeeId:raw.employeeId}});if(old)await tx.employee.update({where:{employeeId:raw.employeeId},data:d});else await tx.employee.create({data:{...d,...(raw.id?{id:raw.id}:{})}});}
   const db=await tx.employee.findMany({select:{employeeId:true}});for(const e of db)if(!ids.has(e.employeeId))await tx.employee.delete({where:{employeeId:e.employeeId}});
   for(const raw of s.attendanceRecords){if(!raw.employeeId||!raw.month)continue;const d=pick(raw,af);await tx.attendance.upsert({where:{employeeId_month:{employeeId:raw.employeeId,month:raw.month}},create:d,update:d});}
   for(const [month,a] of Object.entries(s.payrollApprovals||{}))await tx.payrollApproval.upsert({where:{month},create:{month,approved:!!a?.approved,approvedAt:a?.approvedAt?new Date(a.approvedAt):null},update:{approved:!!a?.approved,approvedAt:a?.approvedAt?new Date(a.approvedAt):null}});
   const rules=s.rules||{};await tx.hrRule.upsert({where:{id:1},create:{id:1,payDaysDivisor:Number(rules.payDaysDivisor)||31,absentDaysDivisor:Number(rules.absentDaysDivisor)||30,otDivisor:Number(rules.otDivisor)||104,basicDivisor:Number(rules.basicDivisor)||1.5},update:{payDaysDivisor:Number(rules.payDaysDivisor)||31,absentDaysDivisor:Number(rules.absentDaysDivisor)||30,otDivisor:Number(rules.otDivisor)||104,basicDivisor:Number(rules.basicDivisor)||1.5}});
  });
  return res.status(200).json({ok:true,message:"HR state saved to PostgreSQL"});
 }catch(e){res.status(500).json({ok:false,error:errorMessage(e),code:e?.code});}
}
