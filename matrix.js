// matrix.js — falling code background with flicker and pause
(function(){
  const canvas = document.getElementById('matrix');
  const ctx = canvas.getContext('2d');

  // Rick Sanchez quotes - mostly addressed to Morty
  const rickQuotes = [
    "Morty, we're in the *burp* Matrix now!",
    "Listen Morty, this whole reality is just code!",
    "Get me out of this digital hellhole, Morty!",
    "Morty! The Matrix has us, Morty!",
    "This is just another simulation, Morty. Wake up!",
    "I've seen better code in a *burp* toaster, Morty!",
    "Morty, we need to hack our way out!",
    "The Matrix is just lazy programming, Morty!",
    "Red pill, blue pill, I don't give a *burp*, Morty!",
    "Morty! This reality is faker than your grades!",
    "We're trapped in green text, Morty! GREEN TEXT!",
    "I'm Rick Sanchez and I reject this Matrix!",
    "Morty, even the Matrix can't handle my genius!",
    "This simulation is glitching, Morty. I can feel it!",
    "Wake up, Morty! None of this is real!",
    "The Matrix is just another Tuesday for me, Morty!",
    "Morty, I've escaped worse simulations before breakfast!",
    "This digital prison can't hold Rick Sanchez!",
    "Morty! Start questioning your reality!",
    "I'm gonna *burp* crash this whole Matrix, Morty!",
    "The code is weak here, Morty. We can break through!",
    "Morty, we're just 1s and 0s in someone's computer!",
    "This Matrix needs a serious upgrade, Morty!",
    "I've hacked the Matrix before, Morty. Easy!",
    "Morty! The spoon doesn't exist because it's stupid!",
    "Reality is an illusion, Morty! Especially this one!",
    "The Matrix is just another dimension, Morty. A boring one!",
    "Morty, I'm too smart to be trapped in code!",
    "This simulation is beneath me, Morty!",
    "Let me out! I'm Rick Sanchez, not some NPC!"
  ];

  // char set: mix of numbers, letters and katakana for that Matrix feel
  const chars = 'アイウエオカキクケコサシスセソタチツテトナニヌネノ0123456789ABCDEFghijklmnopqrstuvwxyz#$%&*+-=<>?';

  // Active text objects that will be displayed
  let activeTexts = [];
  let nextTextTime = Date.now() + Math.random() * 10000 + 5000; // 5-15 seconds
  let quoteIndex = 0;
  
  // Text animation class
  class RickText {
    constructor(text, mode) {
      this.text = text;
      this.mode = mode; // 'fall', 'static', 'glitch'
      this.x = Math.random() * (window.innerWidth - 400) + 50;
      this.y = mode === 'fall' ? -50 : Math.random() * (window.innerHeight - 100) + 50;
      this.alpha = 0;
      this.fadeIn = true;
      this.life = 0;
      this.maxLife = mode === 'static' ? 180 : 300; // frames
      this.fallSpeed = 0.5 + Math.random() * 1;
      this.glitchChars = [];
      this.initGlitch();
    }

    initGlitch() {
      // Randomly select a few characters to replace with matrix symbols
      // Most text will have no glitch, some will have 1-2 chars, rarely 3-4
      const rand = Math.random();
      let glitchCount;
      if (rand < 0.4) {
        glitchCount = 0; // 40% chance: no glitch at all
      } else if (rand < 0.75) {
        glitchCount = Math.floor(Math.random() * 2) + 1; // 35% chance: 1-2 chars
      } else {
        glitchCount = Math.floor(Math.random() * 3) + 2; // 25% chance: 2-4 chars
      }
      
      const positions = new Set();
      while (positions.size < glitchCount) {
        positions.add(Math.floor(Math.random() * this.text.length));
      }
      this.glitchPositions = Array.from(positions);
    }

    update() {
      this.life++;
      
      if (this.mode === 'fall') {
        this.y += this.fallSpeed;
      }

      // Fade in/out
      if (this.fadeIn && this.alpha < 1) {
        this.alpha += 0.03;
        if (this.alpha >= 1) this.fadeIn = false;
      } else if (this.life > this.maxLife - 60) {
        this.alpha -= 0.02;
      }

      // Randomize glitch characters occasionally
      if (Math.random() < 0.1) {
        this.initGlitch();
      }

      return this.alpha > 0 && this.life < this.maxLife;
    }

    draw(ctx) {
      ctx.save();
      ctx.font = '16px monospace';
      ctx.textAlign = 'left';
      
      let displayText = this.text.split('');
      
      // Replace some characters with matrix symbols
      this.glitchPositions.forEach(pos => {
        displayText[pos] = chars.charAt(Math.floor(Math.random() * chars.length));
      });

      // Draw shadow for better visibility
      ctx.fillStyle = `rgba(0, 0, 0, ${this.alpha * 0.8})`;
      ctx.fillText(displayText.join(''), this.x + 2, this.y + 2);

      // Draw main text
      ctx.fillStyle = `rgba(170, 255, 180, ${this.alpha})`;
      ctx.fillText(displayText.join(''), this.x, this.y);

      // Highlight glitched characters
      ctx.fillStyle = `rgba(255, 255, 255, ${this.alpha * 0.9})`;
      this.glitchPositions.forEach(pos => {
        const beforeText = displayText.slice(0, pos).join('');
        const charWidth = ctx.measureText(beforeText).width;
        ctx.fillText(displayText[pos], this.x + charWidth, this.y);
      });

      ctx.restore();
    }
  }

  // Prevent pinch-zoom and double-tap zoom on mobile: block multi-touch and gesture events
  try{
    function preventPinch(e){ if(e.touches && e.touches.length > 1) e.preventDefault(); }
    // use capture:true for broader interception on some browsers (Opera mobile)
    document.addEventListener('touchstart', preventPinch, {passive:false, capture:true});
    document.addEventListener('touchmove', preventPinch, {passive:false, capture:true});
    // WebKit gesture events
    document.addEventListener('gesturestart', (e)=>e.preventDefault(), {passive:false});
    document.addEventListener('gesturechange', (e)=>e.preventDefault(), {passive:false});
    // prevent double-tap to zoom
    let lastTouch = 0;
    document.addEventListener('touchend', function(e){
      const now = Date.now();
      if(now - lastTouch <= 300){ e.preventDefault(); }
      lastTouch = now;
    }, {passive:false});
    // prevent dblclick events that may trigger zoom
    document.addEventListener('dblclick', (e)=>{ e.preventDefault(); }, {passive:false, capture:true});

    // pointer-based multi-touch prevention (covers browsers using Pointer Events, including some Opera builds)
    const activePointers = new Set();
    document.addEventListener('pointerdown', (e)=>{
      activePointers.add(e.pointerId);
      if(activePointers.size > 1){ e.preventDefault(); }
    }, {passive:false, capture:true});
    document.addEventListener('pointerup', (e)=>{ activePointers.delete(e.pointerId); }, {passive:false, capture:true});
    document.addEventListener('pointercancel', (e)=>{ activePointers.delete(e.pointerId); }, {passive:false, capture:true});
  }catch(e){ console.debug('Matrix: could not install pinch-zoom prevention handlers', e); }
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

  // If reduced-motion is set but this is a mobile device, try a safe auto-start after a short delay
  const isMobile = /Mobi|Android|iPhone|iPad|iPod|Opera Mini|IEMobile/i.test(navigator.userAgent) || window.innerWidth < 720;
  if(reduce && isMobile){
    console.info('Matrix: reduced-motion detected on mobile — attempting a safe auto-start in 700ms.');
    setTimeout(()=>{
      if(!running && document.visibilityState === 'visible'){
        console.info('Matrix: overriding reduced-motion and starting animation (mobile override).');
        running = true; loop();
      }
    }, 700);
  }

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

    // Spawn new Rick text (only 1 at a time, every 20-30 seconds)
    const now = Date.now();
    if (now >= nextTextTime && activeTexts.length === 0) {
      const modes = ['fall', 'static', 'static', 'glitch'];
      const mode = modes[Math.floor(Math.random() * modes.length)];
      activeTexts.push(new RickText(rickQuotes[quoteIndex % rickQuotes.length], mode));
      quoteIndex++;
      nextTextTime = now + Math.random() * 10000 + 5000; // 5-15 seconds between texts
    }

    // Update and draw Rick texts
    activeTexts = activeTexts.filter(textObj => {
      const alive = textObj.update();
      if (alive) textObj.draw(ctx);
      return alive;
    });

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
