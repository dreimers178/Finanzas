const fs=require('fs'), sharp=require('sharp');
const topo=require('topojson-client'); const d3=require('d3-geo');
// ---------- Flag of New Zealand (public domain design) ----------
function star(cx,cy,r){const pts=[];for(let i=0;i<10;i++){const a=-Math.PI/2+i*Math.PI/5;const rr=i%2?r*0.382:r;pts.push((cx+rr*Math.cos(a)).toFixed(2)+','+(cy+rr*Math.sin(a)).toFixed(2));}return pts.join(' ');}
const stars=[[900,120,40],[825,270,36],[970,240,40],[900,480,45]];
let s=stars.map(([x,y,r])=>`<polygon points="${star(x,y,r)}" fill="#fff"/><polygon points="${star(x,y,r*0.72)}" fill="#C8102E"/>`).join('');
const flag=`<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="600" viewBox="0 0 1200 600">
<defs><clipPath id="s"><path d="M0,0 v30 h60 v-30 z"/></clipPath>
<clipPath id="t"><path d="M30,15 h30 v15 z v15 h-30 z h-30 v-15 z v-15 h30 z"/></clipPath></defs>
<rect width="1200" height="600" fill="#00247D"/>
<g transform="scale(10)" clip-path="url(#s)">
<path d="M0,0 v30 h60 v-30 z" fill="#00247D"/>
<path d="M0,0 L60,30 M60,0 L0,30" stroke="#fff" stroke-width="6"/>
<path d="M0,0 L60,30 M60,0 L0,30" clip-path="url(#t)" stroke="#C8102E" stroke-width="4"/>
<path d="M30,0 v30 M0,15 h60" stroke="#fff" stroke-width="10"/>
<path d="M30,0 v30 M0,15 h60" stroke="#C8102E" stroke-width="6"/>
</g>${s}</svg>`;
fs.writeFileSync('flag.svg',flag);
// ---------- Map: Pacific view with Mexico and New Zealand ----------
const world=require('world-atlas/countries-50m.json');
const countries=topo.feature(world,world.objects.countries);
const land=topo.feature(world,world.objects.land);
const W=1400,H=1000;
const nz=countries.features.find(f=>f.id==='554'), mx=countries.features.find(f=>f.id==='484'), au=countries.features.find(f=>f.id==='036');
const proj=d3.geoNaturalEarth1().rotate([142,0]);
proj.fitExtent([[40,40],[W-40,H-40]],{type:'FeatureCollection',features:[nz,mx,au,{type:'Feature',geometry:{type:'Point',coordinates:[-140,40]}}]});
const path=d3.geoPath(proj);
const grat=d3.geoGraticule10();
const cdmx=[-99.13,19.43], wlg=[174.78,-41.29];
const km=Math.round(d3.geoDistance(cdmx,wlg)*6371);
console.log('distance km',km);
const line={type:'LineString',coordinates:[cdmx,wlg]};
const [cx,cy]=proj(cdmx),[wx,wy]=proj(wlg);
const [nzx,nzy]=proj([172,-38]);const [mxx,mxy]=proj([-104,25]);
const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
<rect width="${W}" height="${H}" fill="#DCEBF2"/>
<path d="${path(grat)}" fill="none" stroke="#C3DAE5" stroke-width="1.2"/>
<path d="${path(land)}" fill="#C9D3D0" stroke="#AEBBB7" stroke-width="0.8"/>
<path d="${path(mx)}" fill="#C8553D"/>
<path d="${path(nz)}" fill="#2F8F5B" stroke="#0E4D45" stroke-width="1.5"/>
<path d="${path(line)}" fill="none" stroke="#14262B" stroke-width="3" stroke-dasharray="12 10"/>
<circle cx="${cx}" cy="${cy}" r="9" fill="#14262B"/><circle cx="${wx}" cy="${wy}" r="9" fill="#14262B"/>
<circle cx="${nzx}" cy="${nzy}" r="70" fill="none" stroke="#2F8F5B" stroke-width="5"/>
<text x="${nzx+90}" y="${nzy+95}" font-family="Arial" font-weight="bold" font-size="40" fill="#0E4D45" text-anchor="start">NUEVA ZELANDA</text>
<text x="${mxx}" y="${mxy-80}" font-family="Arial" font-weight="bold" font-size="40" fill="#8E3A27" text-anchor="middle">MÉXICO</text>
<text x="${(cx+wx)/2+40}" y="${(cy+wy)/2-20}" font-family="Arial" font-weight="bold" font-size="36" fill="#14262B">≈ 11,100 km</text>
<text x="${proj([-150,-20])[0]}" y="${proj([-150,-20])[1]}" font-family="Arial" font-style="italic" font-size="34" fill="#5F8FA3" text-anchor="middle">Océano Pacífico</text>
<text x="${proj([134,-25])[0]}" y="${proj([134,-25])[1]}" font-family="Arial" font-size="30" fill="#6B7A76" text-anchor="middle">Australia</text>
</svg>`;
fs.writeFileSync('map.svg',svg);
(async()=>{await sharp(Buffer.from(flag)).png().toFile('flag.png');await sharp(Buffer.from(svg)).png().toFile('map.png');console.log('done');})();
