import assert from 'node:assert/strict';
import { test, beforeEach, afterEach } from 'node:test';
import { RequestScheduler, retryAfterMs, throttleScope, newRequestContext, withRequestContext, getScheduler, disposeScheduler, throttleWaitSeconds, coalesceRefresh } from '../src/backend/request-scheduler.ts';
import { graphPolicy, servicePolicy, FOCI_CLIENTS } from '../src/backend/request-policy.ts';
import { GraphClient } from '../src/backend/graph-client.ts';
import { graphFetch, eligible } from '../src/backend/request-transport.ts';
import { acquireFociToken, clearFociTokens, describeToken } from '../src/backend/auth.ts';
import * as tokens from '../src/backend/token-cache.ts';
import { GraphApiError, type PluginAPI } from '../src/shared/types.ts';
import { setLogger } from '../src/backend/logger-singleton.ts';
import { CLIENT_ID_OFFICE as office, CLIENT_ID_TEAMS as teams, GRAPH_SCOPE } from '../src/shared/constants.ts';

const jwt = (client = office, scopes = 'Mail.ReadWrite Mail.Send User.Read', extra = {}) =>
  `e30.${Buffer.from(JSON.stringify({ oid: 'user', tid: 'tenant', appid: client, aud: 'https://graph.microsoft.com', scp: scopes, ...extra })).toString('base64url')}.test`;
const apis: PluginAPI[] = [];
function harness(fetch: typeof globalThis.fetch, prefs: Record<string, unknown> = {}) {
  const api = { fetch, config: { getPluginData: () => ({ preferences: prefs }), setPluginData() {} },
    safeStorage: { isEncryptionAvailable: () => true, encryptString: () => 'encrypted' },
    state: { set() {}, get: () => ({}) } } as unknown as PluginAPI;
  apis.push(api); return api;
}
function seed(scopes?: string) { tokens.set({ accessToken: jwt(office, scopes), refreshToken: 'refresh', expiresAt: Date.now() + 3600_000,
  clientId: office, objectId: 'user', tenantId: 'tenant', scopes: scopes ?? 'Mail.ReadWrite Mail.Send User.Read', email: '', displayName: null }); }
const flush = async () => { for (let i = 0; i < 100; i++) await Promise.resolve(); };
const mail = () => graphPolicy('GET', 'https://graph.microsoft.com/v1.0/me/messages');
const throttle = (scope?: string, after = '10') => new Response(null, { status: 429, headers: { 'Retry-After': after, ...(scope ? { 'x-ms-throttle-scope': scope } : {}) } });
beforeEach(() => { tokens.invalidateSession(); clearFociTokens(); seed(); setLogger({ info() {}, warn() {}, error() {} }); });
afterEach(() => { for (const api of apis.splice(0)) disposeScheduler(api); tokens.invalidateSession(); clearFociTokens(); });

test('retry hints use the longest valid delay and the server clock', () => {
  assert.equal(retryAfterMs(new Headers({ 'rEtRy-AfTeR': '3', 'x-ms-retry-after-ms': '4500', 'retry-after-ms': '2000' })), 4500);
  assert.equal(retryAfterMs(new Headers({ Date: 'Mon, 19 Oct 2026 10:00:00 GMT', 'Retry-After': 'Mon, 19 Oct 2026 10:00:12 GMT' }), 0), 12000);
  assert.equal(retryAfterMs(new Headers({ 'Retry-After': '-1', 'retry-after-ms': 'NaN', 'x-ms-retry-after-ms': '-10' })), undefined);
  assert.equal(retryAfterMs(new Headers({ 'Retry-After': '0' })), 0);
  assert.equal(retryAfterMs(new Headers({ 'Retry-After': 'Mon, 19 Oct 2026 09:00:00 GMT', Date: 'Mon, 19 Oct 2026 10:00:00 GMT' })), 0);
});
test('scope hints must identify this app and tenant', () => {
  assert.equal(throttleScope(new Headers({ 'x-ms-throttle-scope': `Tenant_Application/Read/${office}/tenant` }), office, 'tenant').kind, 'app-tenant');
  assert.equal(throttleScope(new Headers({ 'x-ms-throttle-scope': `Tenant_Application/Read/${teams}/other` }), office, 'tenant').kind, 'unknown');
});
test('shared cooldown survives budget exhaustion and prevents a second client instance from sending', async () => {
  let sends = 0;
  const api = harness(async () => { sends++; return throttle(undefined, '60'); }, { adaptiveFociRouting: false });
  await assert.rejects(new GraphClient(api).listMailFolders(), (e: GraphApiError) => e.statusCode === 429 && e.retry!.retryAt! > Date.now() + 59000);
  const first = sends;
  await assert.rejects(new GraphClient(api).getMail('id'), /Retry in/);
  assert.equal(sends, first);
});
test('429 honors delay; unrelated workloads continue during cooldown', async t => {
  t.mock.timers.enable({ apis: ['Date','setTimeout'] });
  const api = harness(async () => Response.json({})); const s = getScheduler(api); let calls = 0;
  const pending = s.execute(mail(), office, newRequestContext(api), async () => ++calls === 1 ? throttle(undefined,'2') : Response.json({}));
  await flush(); assert.equal(calls, 1);
  await s.execute(servicePolicy('teams-internal','GET'), teams, newRequestContext(api), async () => Response.json({}));
  t.mock.timers.tick(1999); await flush(); assert.equal(calls,1);
  t.mock.timers.tick(1); await pending; assert.equal(calls,2);
});
test('mail concurrency is shared and background leaves one slot', async () => {
  const api = harness(async () => Response.json({})); const s = getScheduler(api);
  const releases: Array<() => void> = []; let active = 0, peak = 0;
  const send = async () => { active++; peak = Math.max(peak,active); await new Promise<void>(r => releases.push(r)); active--; return Response.json({}); };
  const jobs = Array.from({length:4}, () => s.execute(mail(),office,newRequestContext(api,'background'),send));
  await flush(); assert.equal(active,3);
  const interactive = s.execute(mail(),office,newRequestContext(api),send); await flush(); assert.equal(active,4);
  while (releases.length) { releases.shift()!(); await flush(); }
  await Promise.all([...jobs,interactive]); assert.equal(peak,4);
});
test('Teams pacing remains shared across client IDs', async t => {
  t.mock.timers.enable({ apis: ['Date','setTimeout'] });
  const api = harness(async () => Response.json({})); const s = getScheduler(api); let count = 0;
  const p = graphPolicy('GET','https://graph.microsoft.com/v1.0/me/chats');
  await s.execute(p,office,newRequestContext(api),async () => { count++; return Response.json({}); });
  const second = s.execute(p,teams,newRequestContext(api),async () => { count++; return Response.json({}); });
  await flush(); assert.equal(count,1); t.mock.timers.tick(1500); await second; assert.equal(count,2);
});
test('app-specific throttling switches once; tenant throttling never switches', async () => {
  for (const scope of ['Tenant_Application','Tenant']) {
    const api = harness(async () => Response.json({}),{maxThrottleWaitSeconds:0}); const s = getScheduler(api); let alternatives = 0;
    const clients: string[] = [];
    const pending = s.execute(mail(),office,newRequestContext(api),async c => { clients.push(c); return c===office ? throttle(`${scope}/Read/${office}/tenant`) : Response.json({}); },
      { alternate: async () => { alternatives++; return teams; }, refresh: async () => {}, denied() {} });
    if (scope === 'Tenant') { await assert.rejects(pending,/Retry in/); assert.equal(alternatives,0); }
    else { await pending; assert.deepEqual(clients,[office,teams]); assert.equal(alternatives,1); }
  }
});
test('one unknown-scope read probe; correlated throttling blocks all candidates', async () => {
  const api = harness(async () => Response.json({}),{maxThrottleWaitSeconds:0}); const s = getScheduler(api); let alternatives=0;
  await assert.rejects(s.execute(mail(),office,newRequestContext(api),async () => throttle(),
    { alternate: async () => { alternatives++; return teams; },refresh:async()=>{},denied(){} }),/Retry in/);
  assert.equal(alternatives,1);
  assert.ok(s.admissionAt(mail(),FOCI_CLIENTS[3].id)>Date.now());
  let sends=0;
  await assert.rejects(s.execute(mail(),teams,newRequestContext(api),async()=> { sends++; return Response.json({}); }),/Retry in/);
  assert.equal(sends,0);
});
test('unknown writes never probe and ambiguous sends are never repeated', async () => {
  const api = harness(async()=>Response.json({}),{maxThrottleWaitSeconds:0}); const s=getScheduler(api);
  const p=graphPolicy('POST','https://graph.microsoft.com/v1.0/me/sendMail'); let count=0;
  await assert.rejects(s.execute(p,office,newRequestContext(api),async()=> { count++; return throttle(); },
    { alternate:async()=> { throw new Error('must not probe'); },refresh:async()=>{},denied(){} }),/Retry in/);
  assert.equal(count,1);
  const other=harness(async()=>new Response(null,{status:503}));
  await assert.rejects(new GraphClient(other).sendMail({to:[],subject:'test',bodyHtml:''}), (e:GraphApiError)=>!!e.retry?.ambiguous);
});
test('session invalidation cancels queued work without sending', async () => {
  const api=harness(async()=>Response.json({})); const s=getScheduler(api); let count=0;
  const pending=s.execute(mail(),office,newRequestContext(api),async()=> { count++; return throttle(); });
  const rejected=assert.rejects(pending,/session ended/); await flush(); tokens.invalidateSession(); await rejected; assert.equal(count,1);
});
test('scope qualification is explicit and rejects conflicting identity or audience', () => {
  const p=graphPolicy('POST','https://graph.microsoft.com/v1.0/me/sendMail');
  assert.equal(eligible(describeToken(jwt(teams,'Mail.ReadWrite'),teams,GRAPH_SCOPE,Date.now()+10000),p),false);
  assert.equal(eligible(describeToken(jwt(teams,'Mail.Send'),teams,GRAPH_SCOPE,Date.now()+10000),p),true);
  for (const extra of [{aud:'https://outlook.office.com'},{oid:'someone-else'},{tid:'other-tenant'},{appid:office}]) {
    assert.equal(eligible(describeToken(jwt(teams,'Mail.Send',extra),teams,GRAPH_SCOPE,Date.now()+10000),p),false);
  }
  assert.equal(eligible(describeToken('opaque',teams,GRAPH_SCOPE,Date.now()+10000),p),false);
});
test('candidate tokens coalesce and consent failure stays silent and cached', async () => {
  let count=0;
  const api=harness(async()=> { count++; return Response.json({error:'consent_required'},{status:400}); });
  const results=await Promise.allSettled(Array.from({length:8},()=>acquireFociToken(api,teams,GRAPH_SCOPE,{candidate:true})));
  assert.ok(results.every(r=>r.status==='rejected')); assert.equal(count,1);
  await assert.rejects(acquireFociToken(api,teams,GRAPH_SCOPE,{candidate:true})); assert.equal(count,1);
  assert.ok(tokens.getValidAccessToken());
});
test('qualified Graph alternate is used, but routing rollback keeps preferred', async () => {
  for (const enabled of [true,false]) {
    const clients:string[]=[];
    const api=harness(async(input,init)=> {
      if (String(input).includes('login.microsoftonline.com')) {
        const id=new URLSearchParams(String(init?.body)).get('client_id')!;
        return Response.json({access_token:jwt(id,'Mail.Read'),expires_in:3600,token_type:'Bearer'});
      }
      const token=new Headers(init?.headers).get('authorization')!.slice(7);
      const id=JSON.parse(Buffer.from(token.split('.')[1],'base64url').toString()).appid; clients.push(id);
      return id===office ? throttle(`Tenant_Application/Read/${office}/tenant`) : Response.json({value:[]});
    },{adaptiveFociRouting:enabled,maxThrottleWaitSeconds:0});
    const pending=graphFetch(api,'https://graph.microsoft.com/v1.0/me/messages',{headers:{Authorization:`Bearer ${tokens.get()!.accessToken}`}});
    if (enabled) { assert.equal((await pending).status,200); assert.deepEqual(clients,[office,teams]); }
    else { await assert.rejects(pending,/Retry in/); assert.deepEqual(clients,[office]); }
  }
});
test('same-client 401 refresh cannot change application identity', async () => {
  let requests=0;const clients:string[]=[];
  const api=harness(async(input,init)=> {
    if (String(input).includes('login.microsoftonline.com')) { const id=new URLSearchParams(String(init?.body)).get('client_id')!;clients.push(id);return Response.json({access_token:jwt(id),expires_in:3600,token_type:'Bearer'}); }
    return ++requests===1 ? new Response(null,{status:401}):Response.json({id:'user'});
  });
  await new GraphClient(api).getMe();assert.equal(requests,2);assert.deepEqual(clients,[office]);
});
test('token endpoint throttles do not fan out to alternate clients', async () => {
  tokens.set({...tokens.get()!,expiresAt:0});let requests=0;
  const api=harness(async()=> {requests++;return throttle(undefined,'60');});
  await assert.rejects(new GraphClient(api).getMe(),/Retry in/);assert.equal(requests,1);
});
test('configuration preserves zero, bounds values, and snapshots each action', async () => {
  assert.equal(throttleWaitSeconds(undefined),30);assert.equal(throttleWaitSeconds(0),0);assert.equal(throttleWaitSeconds(400),300);assert.equal(throttleWaitSeconds(-1),0);
  const prefs={maxThrottleWaitSeconds:30};const api=harness(async()=>Response.json({}),prefs);
  await withRequestContext(api,async()=> {const before=newRequestContext(api);prefs.maxThrottleWaitSeconds=0;assert.equal(before.remainingMs,30000);assert.equal(newRequestContext(api).remainingMs,0);});
});

test('batch retries only failed items and never mistakes a completed item for a cooldown failure', async t => {
  t.mock.timers.enable({apis:['Date','setTimeout']});
  const batches: string[][]=[]; const retries: string[]=[];
  const api=harness(async(input,init)=> {
    if (init?.method === 'PATCH') { retries.push(String(input)); return new Response(null,{status:204}); }
    if (String(input).includes('/$batch')) {
      const ids=JSON.parse(String(init?.body)).requests.map((r:{id:string})=>r.id);batches.push(ids);
      return Response.json({responses:ids.map((id:string)=>({id,status:batches.length===1 && id==='0' ? 429 : 204,
        ...(id==='0' && batches.length===1 ? {headers:{'retry-after':'2'}}:{})}))});
    }
    return Response.json({value:[{id:'a'},{id:'b'}]});
  },{adaptiveFociRouting:false});
  const pending=new GraphClient(api).markFolderRead('inbox');await flush();assert.equal(batches.length,1);
  t.mock.timers.tick(2000);assert.equal(await pending,2);assert.deepEqual(batches,[['0','1']]); assert.deepEqual(retries,['https://graph.microsoft.com/v1.0/me/messages/a']);
});
test('partial batch errors count successful peers even after retry budget expires', async () => {
  const api=harness(async(input)=>String(input).includes('/$batch') ? Response.json({responses:[
    {id:'0',status:429,headers:{'Retry-After':'60'}},{id:'1',status:204},
  ]}) : Response.json({value:[{id:'a'},{id:'b'}]}),{maxThrottleWaitSeconds:0,adaptiveFociRouting:false});
  await assert.rejects(new GraphClient(api).markFolderRead('inbox'),(e:GraphApiError)=>{
    assert.deepEqual(e.retry?.partialResult,{succeeded:1,failed:1});assert.ok(e.retry?.retryAt);return true;
  });
});
test('batch reservations cannot overlap other traffic for the same mailbox', async () => {
  const api=harness(async()=>Response.json({}));const s=getScheduler(api);let release!:()=>void;let batchStarted=false;
  const ordinary=s.execute(mail(),office,newRequestContext(api),async()=>{await new Promise<void>(r=>release=r);return Response.json({});});
  await flush();const batch=s.execute(graphPolicy('POST','https://graph.microsoft.com/v1.0/$batch'),office,newRequestContext(api),async()=>{batchStarted=true;return Response.json({});});
  await flush();assert.equal(batchStarted,false);release();await Promise.all([ordinary,batch]);assert.equal(batchStarted,true);
});
test('uncertain upload chunk checks progress before advancing and sends no Graph token to upload URL', async t => {
  t.mock.timers.enable({apis:['Date','setTimeout']});const methods:string[]=[];
  const api=harness(async(input,init)=>{
    if (String(input).includes('graph.microsoft.com')) return Response.json({uploadUrl:'https://upload.example.test/session'});
    assert.equal(new Headers(init?.headers).has('authorization'),false);
    methods.push(init?.method ?? 'GET');
    if (methods.length===1) return new Response(null,{status:503,headers:{'Retry-After':'1'}});
    return Response.json({id:'uploaded',name:'file',webUrl:'https://example.test/file'});
  });
  const pending=new GraphClient(api).uploadDriveFile('file',Buffer.alloc(5*1024*1024),'application/octet-stream');
  await flush();t.mock.timers.tick(1000);const item=await pending;
  assert.equal(item.id,'uploaded');assert.deepEqual(methods,['PUT','GET']);
});
test('expired background refreshes coalesce and resume once at the server deadline', async t => {
  t.mock.timers.enable({apis:['Date','setTimeout']});let calls=0;
  const api=harness(async()=>Response.json({}),{maxThrottleWaitSeconds:0});
  const fn=async()=>{if(++calls===1) throw new GraphApiError('later',429,undefined,{retryAt:Date.now()+5000});return 42;};
  const a=coalesceRefresh(api,'poll',fn),b=coalesceRefresh(api,'poll',fn);assert.equal(a,b);
  await flush();t.mock.timers.tick(4999);await flush();assert.equal(calls,1);
  t.mock.timers.tick(1);assert.equal(await a,42);assert.equal(calls,2);
});
test('four attempts cap includes refresh and prevents endless retries', async t => {
  t.mock.timers.enable({apis:['Date','setTimeout']});const api=harness(async()=>Response.json({}));let calls=0;
  const pending=getScheduler(api).execute(mail(),office,newRequestContext(api),async()=>{calls++;return calls===1?new Response(null,{status:401}):throttle(undefined,'1');},
    {alternate:async()=>undefined,refresh:async()=>{},denied(){}});
  const rejected=assert.rejects(pending,(e:GraphApiError)=>e.retry?.attempts===4);
  for(let i=0;i<4;i++){await flush();t.mock.timers.tick(1000);}await rejected;assert.equal(calls,4);
});
test('a previously successful unknown-scope alternate cannot become a chain of probes', async () => {
  const api=harness(async()=>Response.json({}),{maxThrottleWaitSeconds:0});const s=getScheduler(api);let alternatives=0;
  const routes={alternate:async()=>{alternatives++;return teams;},refresh:async()=>{},denied(){}};
  await s.execute(mail(),office,newRequestContext(api),async client=>client===office?throttle():Response.json({}),routes);
  await assert.rejects(s.execute(mail(),office,newRequestContext(api),async()=>throttle(),routes),/Retry in/);
  assert.ok(s.admissionAt(mail(),FOCI_CLIENTS[3].id)>Date.now());
  assert.equal(alternatives,2); // One original probe, then reuse of the temporary hint.
  await assert.rejects(s.execute(mail(),office,newRequestContext(api),async()=>{throw new Error('must not send');},routes),/Retry in/);
  assert.equal(alternatives,2);
});
test('FOCI candidate validation is serialized at one request per second', async t => {
  t.mock.timers.enable({apis:['Date','setTimeout']});let calls=0;
  const api=harness(async(_input,init)=> {calls++;const id=new URLSearchParams(String(init?.body)).get('client_id')!;return Response.json({access_token:jwt(id,'Mail.Read'),expires_in:3600,token_type:'Bearer'});});
  const first=acquireFociToken(api,teams,GRAPH_SCOPE,{candidate:true});
  const second=acquireFociToken(api,FOCI_CLIENTS[2].id,GRAPH_SCOPE,{candidate:true});
  await first;await flush();assert.equal(calls,1);
  t.mock.timers.tick(1000);await second;assert.equal(calls,2);
});
test('a forbidden alternate is authoritative and is not followed by another client', async () => {
  const clients:string[]=[];
  const api=harness(async(input,init)=> {
    if(String(input).includes('login.microsoftonline.com')) {const id=new URLSearchParams(String(init?.body)).get('client_id')!;return Response.json({access_token:jwt(id,'Mail.Read'),expires_in:3600,token_type:'Bearer'});}
    const id=JSON.parse(Buffer.from(new Headers(init?.headers).get('authorization')!.slice(7).split('.')[1],'base64url').toString()).appid;
    clients.push(id);return id===office?throttle(`Application/Read/${office}/tenant`):new Response(null,{status:403});
  },{maxThrottleWaitSeconds:0});
  const response=await graphFetch(api,'https://graph.microsoft.com/v1.0/me/messages',{headers:{Authorization:`Bearer ${tokens.get()!.accessToken}`}});
  assert.equal(response.status,403);assert.deepEqual(clients,[office,teams]);
});
test('batch envelope retries count toward each item attempt limit', async t => {
  t.mock.timers.enable({apis:['Date','setTimeout']});let batches=0;
  const api=harness(async input=>{
    if(!String(input).includes('/$batch')) return Response.json({value:[{id:'a'},{id:'b'}]});
    if(++batches<4) return throttle(undefined,'1');
    return Response.json({responses:[{id:'0',status:429,headers:{'Retry-After':'1'}},{id:'1',status:204}]});
  },{adaptiveFociRouting:false});
  const rejected=assert.rejects(new GraphClient(api).markFolderRead('inbox'),(e:GraphApiError)=>{
    assert.deepEqual(e.retry?.partialResult,{succeeded:1,failed:1});assert.equal(e.retry?.attempts,4);return true;
  });
  for(let i=0;i<4;i++){await flush();t.mock.timers.tick(1000);}await rejected;assert.equal(batches,4);
});
test('an app-throttled batch item can use a scope-qualified alternate without replaying its peers', async () => {
  const writes:string[]=[];
  const api=harness(async(input,init)=>{
    if(String(input).includes('login.microsoftonline.com')) return Response.json({access_token:jwt(teams,'Mail.ReadWrite'),expires_in:3600,token_type:'Bearer'});
    if(String(input).includes('/$batch')) return Response.json({responses:[{id:'0',status:429,headers:{'Retry-After':'10','x-ms-throttle-scope':`Tenant_Application/Write/${office}/tenant`}}, {id:'1',status:204}]});
    if(init?.method==='PATCH') {writes.push(new Headers(init.headers).get('authorization')!);return new Response(null,{status:204});}
    return Response.json({value:[{id:'a'},{id:'b'}]});
  },{maxThrottleWaitSeconds:0});
  assert.equal(await new GraphClient(api).markFolderRead('inbox'),2);
  assert.deepEqual(writes,[`Bearer ${jwt(teams,'Mail.ReadWrite')}`]);
});
