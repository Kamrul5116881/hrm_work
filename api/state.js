import { prisma } from "./_lib/prisma.js";
import { methodGuard, body, errorMessage } from "./_lib/http.js";
import { requireAuth, canWrite } from "./_lib/auth.js";
const ef=["employeeId","name","joiningDate","jobTitle","section","basic","houseRent","conveyance","food","medical","gross","casualLeaveAlloc","medicalLeaveAlloc","motherName","fatherName","dob","nid","maritalStatus","status"];
const af=["employeeId","month","present","weekend","leave","absent","otHours","advance","arrear","tds","basic","medical","conveyance","food","gross","totalDays","payableDays"];
const nums=new Set(["basic","houseRent","conveyance","food","medical","gross","casualLeaveAlloc","medicalLeaveAlloc","present","weekend","leave","absent","otHours","advance","arrear","tds","totalDays","payableDays"]);
// Per-month salary overrides are nullable in the schema: absent/blank means
// "fall back to the employee master value". They must NEVER be coerced to 0,
// or every reload replaces real salaries with zeros.
const nullableSalary=new Set(["basic","medical","conveyance","food","gross"]);
function pick(x,fs,nullableOk){const d={};for(const k of fs){if(x[k]===undefined)continue;if(nullableOk&&nullableSalary.has(k)){d[k]=x[k]===null||x[k]===""?null:Number(x[k])||0;}else d[k]=nums.has(k)?Number(x[k])||0:x[k];}return d;}
export default async function handler(req,res){
 if(!methodGuard(req,res,["GET","POST"]))return;
 // Session-token auth (Bearer). 401 for missing/invalid/expired tokens.
 const auth = requireAuth(req,res);
 if (auth !== true) return;
 try{
  if(req.method==="POST" && !canWrite(req.auth.role)){
   return res.status(403).json({success:false,error:"Write access denied for your role. HR/Admin roles only."});
  }
   if(req.method==="GET"){
    const [employees,attendanceRecords,payrollRecords,approvals,rules]=await Promise.all([prisma.employee.findMany({orderBy:{employeeId:"asc"}}),prisma.attendance.findMany({orderBy:[{month:"desc"},{employeeId:"asc"}]}),prisma.payroll.findMany({orderBy:[{month:"desc"},{employeeId:"asc"}]}),prisma.payrollApproval.findMany(),prisma.hrRule.findUnique({where:{id:1}})]);
    return res.status(200).json({state:{employees:employees.map(({createdAt,updatedAt,...e})=>e),attendanceRecords:attendanceRecords.map(({createdAt,updatedAt,...a})=>a),payrollRecords:payrollRecords.map(({createdAt,updatedAt,...p})=>p),payrollApprovals:Object.fromEntries(approvals.map(a=>[a.month,{approved:a.approved,approvedAt:a.approvedAt?.toISOString()||null}])),rules:rules||{payDaysDivisor:31,absentDaysDivisor:30,otDivisor:104,basicDivisor:1.5}}});
  }
   const b=body(req),s=b.state||b;if(!s||!Array.isArray(s.employees)||!Array.isArray(s.attendanceRecords))return res.status(400).json({error:"Expected state.employees and state.attendanceRecords"});
   await prisma.$transaction(async tx=>{
    // Employees: bulk-create new rows, per-row update for existing (state-sync semantics:
    // employees absent from the payload are removed below).
    const ids=new Set(s.employees.map(e=>e.employeeId).filter(Boolean));
    const existing=new Set((await tx.employee.findMany({select:{employeeId:true}})).map(e=>e.employeeId));
    const creates=[],updates=[];
    for(const raw of s.employees){if(!raw.employeeId||!raw.name)continue;const d=pick(raw,ef);(existing.has(raw.employeeId)?updates:creates).push({...d,employeeId:raw.employeeId});}
    if(creates.length)await tx.employee.createMany({data:creates});
    for(const u of updates)await tx.employee.update({where:{employeeId:u.employeeId},data:u});
    const db=await tx.employee.findMany({select:{employeeId:true}});for(const e of db)if(!ids.has(e.employeeId))await tx.employee.delete({where:{employeeId:e.employeeId}});
    // Attendance: replace-by-month in bulk (fixes broken upsert that omitted employeeId/month on create).
    const byMonth={};
    for(const raw of s.attendanceRecords){if(!raw.employeeId||!raw.month)continue;(byMonth[raw.month]??=[]).push({...pick(raw,af,true),employeeId:raw.employeeId,month:raw.month});}
    for(const [month,rows] of Object.entries(byMonth)){await tx.attendance.deleteMany({where:{month}});if(rows.length)await tx.attendance.createMany({data:rows});}
    for(const [month,a] of Object.entries(s.payrollApprovals||{}))await tx.payrollApproval.upsert({where:{month},create:{month,approved:!!a?.approved,approvedAt:a?.approvedAt?new Date(a.approvedAt):null},update:{approved:!!a?.approved,approvedAt:a?.approvedAt?new Date(a.approvedAt):null}});
    const rules=s.rules||{};await tx.hrRule.upsert({where:{id:1},create:{id:1,payDaysDivisor:Number(rules.payDaysDivisor)||31,absentDaysDivisor:Number(rules.absentDaysDivisor)||30,otDivisor:Number(rules.otDivisor)||104,basicDivisor:Number(rules.basicDivisor)||1.5},update:{payDaysDivisor:Number(rules.payDaysDivisor)||31,absentDaysDivisor:Number(rules.absentDaysDivisor)||30,otDivisor:Number(rules.otDivisor)||104,basicDivisor:Number(rules.basicDivisor)||1.5}});
    // Payroll ledger: derived SERVER-SIDE from employees x attendance x rules
    // after every save — records salary rows for EVERY month already in the
    // database, no matter which client or app version triggered the write.
    // Mirrors computePayroll() in src/App.jsx exactly.
    const [emps,atts,apprRows,hrrule]=await Promise.all([tx.employee.findMany(),tx.attendance.findMany(),tx.payrollApproval.findMany(),tx.hrRule.findUnique({where:{id:1}})]);
    const R=hrrule||{payDaysDivisor:31,absentDaysDivisor:30,otDivisor:104,basicDivisor:1.5};
    const empById=new Map(emps.map(e=>[e.employeeId,e])),approvedMonths=new Set(apprRows.filter(a=>a.approved).map(a=>a.month));
    const nn=(v)=>{const x=parseFloat(v);return isNaN(x)?0:x;},r2=(x)=>Math.round((x+Number.EPSILON)*100)/100;
    const prows=[];
    for(const att of atts){
     const emp=empById.get(att.employeeId);if(!emp||!att.month)continue;
     const present=nn(att.present),weekend=nn(att.weekend),leave=nn(att.leave),absent=nn(att.absent),otHours=nn(att.otHours),advance=nn(att.advance),arrear=nn(att.arrear),tds=nn(att.tds);
     const medical=nn(att.medical??emp.medical),conveyance=nn(att.conveyance??emp.conveyance),food=nn(att.food??emp.food),gross=nn(att.gross??emp.gross);
     const totalDays=present+weekend+leave+absent,payableDays=present+weekend+leave;
     const basic=(gross-(medical+conveyance+food))/R.basicDivisor,houseRent=basic*0.5;
     const paySalary=(gross/R.payDaysDivisor)*totalDays,absentAmount=(basic/R.absentDaysDivisor)*absent,otRate=basic/R.otDivisor,otAmount=otRate*otHours;
     const actualAmount=paySalary-absentAmount+otAmount,payBeforeTds=actualAmount-advance+arrear,payAmount=payBeforeTds-tds;
     prows.push({employeeId:att.employeeId,month:att.month,status:approvedMonths.has(att.month)?"Approved":"Draft",basic:r2(basic),houseRent:r2(houseRent),medical:r2(medical),conveyance:r2(conveyance),food:r2(food),gross:r2(gross),totalDays,payableDays,present,weekend,leave,absent,paySalary:r2(paySalary),absentAmount:r2(absentAmount),otRate:r2(otRate),otHours,otAmount:r2(otAmount),actualAmount:r2(actualAmount),advance,arrear,payBeforeTds:r2(payBeforeTds),tds,payAmount:r2(payAmount)});
    }
    await tx.payroll.deleteMany({});
    if(prows.length)await tx.payroll.createMany({data:prows});
   },{maxWait:10000,timeout:120000});
  return res.status(200).json({ok:true,message:"HR state saved to PostgreSQL"});
 }catch(e){res.status(500).json({ok:false,error:errorMessage(e),code:e?.code});}
}
