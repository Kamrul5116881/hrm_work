import { prisma } from "./_lib/prisma.js";
import { methodGuard, body, errorMessage } from "./_lib/http.js";
import { requireAuth, canWrite } from "./_lib/auth.js";

const TEXT = ["date","cheque","chequeDate","vendor","paymentType","particular","vendorAddress","tin","bin","mainGL","subGL","sectionRef","condition","remarks"];
const NUMS = ["invoiceAmount","tdsRate","vdsRate"];
function pick(x){
  const d={};
  for(const k of TEXT)if(x[k]!==undefined)d[k]=String(x[k]);
  for(const k of NUMS){const n=Number(x[k]);if(x[k]!==undefined&&x[k]!==""&&!Number.isNaN(n))d[k]=n;}
  return d;
}

export default async function handler(req,res){
 if(!methodGuard(req,res,["GET","POST"]))return;
 const auth=requireAuth(req,res);
 if(auth!==true)return;
 try{
  if(req.method==="GET"){
   const rows=await prisma.ledgerTransaction.findMany({orderBy:{createdAt:"asc"}});
   return res.status(200).json({records:rows.map(({createdAt,updatedAt,...r})=>r)});
  }
  if(!canWrite(req.auth.role)){
   return res.status(403).json({success:false,error:"Write access denied for your role."});
  }
  const b=body(req),records=b.records;
  if(!Array.isArray(records))return res.status(400).json({error:"Expected records array"});
  await prisma.$transaction(async tx=>{
   const ids=new Set(records.map(r=>r.id).filter(Boolean));
   const existing=new Set((await tx.ledgerTransaction.findMany({select:{id:true}})).map(r=>r.id));
   const creates=[],updates=[];
   for(const raw of records){
    const d=pick(raw),id=raw.id;
    if(id&&existing.has(id)){const {id:_i,...data}=d;updates.push({where:{id},data});}
    else{creates.push(id?{id,...d}:d);}
   }
   if(creates.length)await tx.ledgerTransaction.createMany({data:creates});
   for(const u of updates)await tx.ledgerTransaction.update(u);
   const db=await tx.ledgerTransaction.findMany({select:{id:true}});
   for(const r of db)if(!ids.has(r.id))await tx.ledgerTransaction.delete({where:{id:r.id}});
  },{maxWait:10000,timeout:60000});
  return res.status(200).json({ok:true,message:"Ledger saved to PostgreSQL"});
 }catch(e){res.status(500).json({ok:false,error:errorMessage(e),code:e?.code});}
}
