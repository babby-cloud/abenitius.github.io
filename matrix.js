// matrix.js — falling code background with flicker and pause
(function(){
  const canvas = document.getElementById('matrix');
  const ctx = canvas.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  // CSS pixel logical size
  let W = window.innerWidth;
  let H = window.innerHeight;

  function setupCanvas(){
    W = Math.max(1, window.innerWidth);
    H = Math.max(1, window.innerHeight);
    // set actual drawing buffer to device pixels
    canvas.width = Math.max(1, Math.floor(W * dpr));
    canvas.height = Math.max(1, Math.floor(H * dpr));
    canvas.style.width = W + 'px';
    canvas.style.height = H + 'px';
    // reset any transform and scale once to match device pixels
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  const alphaFade = 0.05; // background alpha for trails
  // char set: mix of numbers, letters and katakana for that Matrix feel
  const chars = 'アイウエオカキクケコサシスセソタチツテトナニヌネノ0123456789ABCDEFghijklmnopqrstuvwxyz#$%&*+-=<>?';
  let fontSize = Math.max(12, Math.floor(Math.min(W, H) / 60));
  let columns = Math.max(2, Math.floor(W / fontSize) + 1);
  let drops = new Array(columns).fill(0).map(()=>Math.floor(Math.random()*H/fontSize));

  function initSizes(){
    fontSize = Math.max(12, Math.floor(Math.min(window.innerWidth, window.innerHeight) / 60));
    ctx.font = fontSize + 'px monospace';
    columns = Math.max(2, Math.floor(window.innerWidth / fontSize) + 1);
    drops = new Array(columns).fill(0).map(()=>Math.floor(Math.random()*window.innerHeight/fontSize));
  }

  let running = true;
  let allowFlick = true;

  window.addEventListener('resize', ()=>{
    // reconfigure canvas buffer and re-calc sizes
    setupCanvas();
    initSizes();
    // draw a frame so the user sees immediate content
    draw();
  });

  // prefer reduced motion: start paused but draw a frame so the canvas isn't blank
  const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  if(reduce){ running = false; console.info('Matrix: prefers-reduced-motion is set; starting paused.'); }

  // init canvas and sizes
  setupCanvas();
  initSizes();

  function draw(){
    // translucent black to create the trailing effect
    ctx.fillStyle = 'rgba(0,0,0,'+alphaFade+')';
    ctx.fillRect(0,0,Math.max(1, window.innerWidth), Math.max(1, window.innerHeight));

    ctx.textAlign = 'left';
    for(let i=0;i<columns;i++){
      const x = i * fontSize;
      const y = drops[i] * fontSize;
      const r = Math.floor(Math.random()*chars.length);
      const text = chars.charAt(r);

      // head in bright green
      ctx.fillStyle = 'rgba(170,255,180,0.95)';
      ctx.fillText(text, x, y);

      // small chance to create brighter streaks
      if(Math.random() > 0.96){ ctx.fillStyle = 'rgba(200,255,190,0.9)'; ctx.fillText(text, x, y - fontSize); }

      drops[i]++;
      if(drops[i] * fontSize > window.innerHeight && Math.random() > 0.975){ drops[i] = 0; }
    }

    // occasional subtle flicker (global brightness pulse)
    if(allowFlick && Math.random() < 0.002){
      const card = document.querySelector('.card');
      if(card){ card.classList.add('flick'); setTimeout(()=>card.classList.remove('flick'), 140); }
    }
  }

  // animation loop
  let rafId = null;
  function loop(){ if(!running) return; draw(); rafId = requestAnimationFrame(loop); }
  // draw one frame so the canvas isn't blank when paused initially
  try{ draw(); }catch(e){ console.error('Matrix: draw() failed', e); }

  if(running) loop();

  // controls
  const togglePause = document.getElementById('togglePause');
  const toggleFlick = document.getElementById('toggleFlick');

  function setPaused(val){
    running = !val;
    if(!running){ cancelAnimationFrame(rafId); } else { loop(); }
    if(typeof togglePause !== 'undefined' && togglePause) try{ togglePause.textContent = running ? 'Pause' : 'Resume'; }catch(e){}
  }

  function setFlick(val){
    allowFlick = val;
    if(typeof toggleFlick !== 'undefined' && toggleFlick) try{ toggleFlick.textContent = 'Flicker: ' + (allowFlick ? 'On' : 'Off'); }catch(e){}
  }

  if(togglePause) togglePause.addEventListener('click', ()=> setPaused(running));
  if(toggleFlick) toggleFlick.addEventListener('click', ()=> setFlick(!allowFlick));

  // Auto-start after a short timeout on mobile/slow devices (unless user prefers reduced motion)
  if(!running && !reduce){
    // try to start a couple times; some mobile browsers take time to layout/paint
    setTimeout(()=>{ if(!running && document.visibilityState === 'visible'){ running = true; loop(); } }, 250);
    setTimeout(()=>{ if(!running && document.visibilityState === 'visible'){ running = true; loop(); } }, 1000);
    // also start when tab becomes visible
    document.addEventListener('visibilitychange', ()=>{ if(document.visibilityState === 'visible' && !running){ running = true; loop(); } });
  }

  // keyboard shortcuts
  window.addEventListener('keydown', (e)=>{ if(e.code === 'Space'){ e.preventDefault(); setPaused(running); } });

  // expose small API for dev console
  window._matrix = { pause: ()=>setPaused(true), resume: ()=>setPaused(false), flickOn: ()=>setFlick(true), flickOff: ()=>setFlick(false) };
})();
