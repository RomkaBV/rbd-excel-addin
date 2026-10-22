import express from "express";
import helmet from "helmet";
import cors from "cors";
import {createRemoteJWKSet, jwtVerify} from "jose";

const app=express();
app.disable("x-powered-by");
app.use(helmet());
app.use(express.json({limit:"16kb"}));

const tenant=process.env.ENTRA_TENANT_ID;
const audience=process.env.ENTRA_API_AUDIENCE;
const origin=process.env.RBD_ORIGIN;
const port=Number(process.env.PORT||8080);
if(!tenant||!audience||!origin||!/^https:\/\//.test(origin)){
  throw new Error("Configure ENTRA_TENANT_ID, ENTRA_API_AUDIENCE, RBD_ORIGIN");
}
const issuer="https://login.microsoftonline.com/"+tenant+"/v2.0";
const jwks=createRemoteJWKSet(new URL("https://login.microsoftonline.com/"+tenant+"/discovery/v2.0/keys"));
app.use(cors({origin,methods:["GET","POST"],allowedHeaders:["Authorization","Content-Type","Idempotency-Key"],optionsSuccessStatus:204}));

async function authenticate(req,res,next){
  try{
    const match=/^Bearer (.+)$/i.exec(req.headers.authorization||"");
    if(!match)return res.status(401).json({error:"Authorization required"});
    const {payload}=await jwtVerify(match[1],jwks,{issuer,audience,algorithms:["RS256"]});
    if(payload.tid!==tenant||!payload.oid)return res.status(401).json({error:"Invalid identity"});
    req.actor={oid:payload.oid,tenant:payload.tid};
    next();
  }catch(_){res.status(401).json({error:"Invalid access token"});}
}
app.get("/health",(_,res)=>res.json({status:"ok",googleConnected:false}));
app.use("/api",authenticate);
app.param("requestId",(req,res,next,value)=>{
  if(!/^[\p{L}\p{N}_-]{3,80}$/u.test(value))return res.status(400).json({error:"Invalid request ID"});
  next();
});
async function authorizeRequest(req,res,next){
  // Rights must be checked against authoritative server-side ACL; never Excel worksheet names.
  // FAIL CLOSED while the ACL and storage integration are unconfigured.
  return res.status(503).json({error:"Request ACL backend not configured. Access denied."});
}
app.use("/api/requests/:requestId",authorizeRequest);
app.post("/api/requests/:requestId/letters",(_,res)=>res.sendStatus(503));
app.get("/api/requests/:requestId/letters",(_,res)=>res.sendStatus(503));
app.get("/api/requests/:requestId/letters/:fileId",(_,res)=>res.sendStatus(503));
app.post("/api/requests/:requestId/reminders",(_,res)=>res.sendStatus(503));
app.listen(port,()=>console.log("RBD API listening on",port));
