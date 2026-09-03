// window.claude.use("db") の代わりに、同じ契約のインメモリ実装を入れる
(function(){
  var docs = {};            // path -> body
  var subs = [];            // {kind, path, cb}
  function clone(o){ return JSON.parse(JSON.stringify(o)); }
  function fire(){
    subs.slice().forEach(function(s){
      if(s.kind === "doc"){
        var d = docs[s.path];
        s.cb({ id: s.path.split("/").pop(), exists: !!d,
               data: function(){ return d ? clone(d) : undefined; },
               metadata:{fromCache:false, hasPendingWrites:false} });
      } else {
        var out = [];
        Object.keys(docs).forEach(function(p){
          var parent = p.slice(0, p.lastIndexOf("/"));
          if(parent === s.path) out.push({ id: p.split("/").pop(), path: p });
        });
        if(s.order) out.sort(function(a,b){
          var av = docs[a.path][s.order.f], bv = docs[b.path][s.order.f];
          return (s.order.d === "desc" ? -1 : 1) * (av < bv ? -1 : av > bv ? 1 : 0);
        });
        if(s.lim) out = out.slice(0, s.lim);
        s.cb({ docs: out.map(function(x){
                 return { id:x.id, exists:true, data:function(){ return clone(docs[x.path]); },
                          metadata:{fromCache:false,hasPendingWrites:false} };
               }), size: out.length, empty: !out.length,
               docChanges: function(){ return []; },
               metadata:{fromCache:false, hasPendingWrites:false} });
      }
    });
  }
  function docRef(path){
    if(path.split("/").length % 2) throw new TypeError("doc path must have an even number of segments: " + path);
    return {
      id: path.split("/").pop(), path: path,
      get: function(){ var d = docs[path];
        return Promise.resolve({ id:this.id, exists:!!d, data:function(){ return d?clone(d):undefined; },
                                 metadata:{fromCache:false,hasPendingWrites:false} }); },
      set: function(v){ docs[path] = clone(v); fire(); return Promise.resolve(); },
      update: function(v){
        if(!docs[path]) return Promise.reject({code:"invalid_argument", message:"no such document"});
        Object.keys(v).forEach(function(k){ docs[path][k] = clone(v[k]); });
        fire(); return Promise.resolve(); },
      delete: function(){ delete docs[path]; fire(); return Promise.resolve(); },
      onSnapshot: function(cb){ var s={kind:"doc",path:path,cb:cb}; subs.push(s); setTimeout(fire,0);
        return function(){ subs = subs.filter(function(x){ return x!==s; }); }; },
      collection: function(p){ return collRef(path + "/" + p); }
    };
  }
  function collRef(path, order, lim){
    if(path.split("/").length % 2 === 0) throw new TypeError("collection path must have an odd number of segments: " + path);
    return {
      path: path,
      doc: function(id){ return docRef(path + "/" + (id || ("auto" + Math.random().toString(36).slice(2)))); },
      add: function(v){ var r = this.doc(); return r.set(v).then(function(){ return r; }); },
      orderBy: function(f,d){ return collRef(path, {f:f,d:d}, lim); },
      limit: function(n){ return collRef(path, order, n); },
      where: function(){ return this; },
      get: function(){ return Promise.resolve({docs:[],size:0,empty:true,docChanges:function(){return[]},
                       metadata:{fromCache:false,hasPendingWrites:false}}); },
      onSnapshot: function(cb){ var s={kind:"coll",path:path,cb:cb,order:order,lim:lim}; subs.push(s);
        setTimeout(fire,0); return function(){ subs = subs.filter(function(x){ return x!==s; }); }; }
    };
  }
  window.__dump = function(){ return clone(docs); };
  window.claude = { use: function(n){ return Promise.resolve(n === "db" ? { doc: docRef, collection: collRef } : null); } };
})();
