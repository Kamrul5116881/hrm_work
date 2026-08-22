import { prisma } from "./_lib/prisma.js";
import { methodGuard, body, errorMessage } from "./_lib/http.js";
export default async function handler(req,res){
 if(!methodGuard(req,res,["GET","POST","PUT","PATCH","DELETE"]))return;
 try{
  if(req.method==="GET")return res.status(200).json({requests:await prisma.leaveRequest.findMany({where:req.query?.employeeId?{employeeId:req.query.employeeId}:undefined,orderBy:{createdAt:"desc"}})});
  const b=body(req);if(req.method==="POST"){if(!b.employeeId||!b.leaveType)return res.status(400).json({error:"employeeId and leaveType are required"});return res.status(201).json({request:await prisma.leaveRequest.create({data:{employeeId:b.employeeId,leaveType:b.leaveType,startDate:b.startDate||"",endDate:b.endDate||"",totalDays:Number(b.totalDays)||0,reason:b.reason||null,status:b.status||"Pending"}})});}
  if(!b.id)return res.status(400).json({error:"id is required"});
  if(req.method==="DELETE"){await prisma.leaveRequest.delete({where:{id:b.id}});return res.status(200).json({ok:true});}
  return res.status(200).json({request:await prisma.leaveRequest.update({where:{id:b.id},data:{leaveType:b.leaveType,startDate:b.startDate||"",endDate:b.endDate||"",totalDays:Number(b.totalDays)||0,reason:b.reason||null,status:b.status||"Pending",approvedBy:b.approvedBy||null,approvedAt:b.status==="Approved"?new Date():null}})});
 }catch(e){res.status(500).json({error:errorMessage(e),code:e?.code});}
}
