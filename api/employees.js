import { prisma } from "./_lib/prisma.js";
import { methodGuard, body, errorMessage } from "./_lib/http.js";
const fields=["employeeId","name","joiningDate","jobTitle","section","basic","houseRent","conveyance","food","medical","gross","casualLeaveAlloc","medicalLeaveAlloc","motherName","fatherName","dob","nid","maritalStatus","status"];
const nums=new Set(["basic","houseRent","conveyance","food","medical","gross","casualLeaveAlloc","medicalLeaveAlloc"]);
function data(x){const d={}; for(const k of fields) if(x[k]!==undefined)d[k]=nums.has(k)?Number(x[k])||0:x[k]; return d;}
export default async function handler(req,res){
 if(!methodGuard(req,res,["GET","POST","PUT","PATCH","DELETE"]))return;
 try{
  if(req.method==="GET"){const id=req.query?.employeeId; const employees=id?[await prisma.employee.findUnique({where:{employeeId:id}})].filter(Boolean):await prisma.employee.findMany({orderBy:{employeeId:"asc"}}); return res.status(200).json({employees});}
  const b=body(req), id=b.employeeId||req.query?.employeeId;
  if(req.method==="POST"){if(!b.employeeId||!b.name)return res.status(400).json({error:"employeeId and name are required"}); return res.status(201).json({employee:await prisma.employee.create({data:{...data(b),...(b.id?{id:b.id}:{})}})});}
  if(!id)return res.status(400).json({error:"employeeId is required"});
  if(req.method==="DELETE"){await prisma.employee.delete({where:{employeeId:id}});return res.status(200).json({ok:true});}
  return res.status(200).json({employee:await prisma.employee.update({where:{employeeId:id},data:data(b)})});
 }catch(e){res.status(e?.code==="P2002"?409:e?.code==="P2025"?404:500).json({error:errorMessage(e),code:e?.code});}
}
