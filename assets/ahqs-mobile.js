(function(){
  // Mark the current page in the navigation
  try{
    var path=location.pathname.replace(/\/+$/,'/');var file=path.split('/').pop()||'index.html';
    document.querySelectorAll('.ahqs-primary-nav a,.ahqs-mobile-links a').forEach(function(a){
      if(a.classList.contains('ahqs-nav-cta'))return;
      var h=(a.getAttribute('href')||'').split('#')[0];if(h==='/'||h==='')h='index.html';
      if(h.split('/').pop()===file)a.setAttribute('aria-current','page');
    });
  }catch(e){}
  // Close the mobile menu after a link is tapped or when tapping outside
  document.addEventListener('click',function(e){
    document.querySelectorAll('details.ahqs-mobile-menu[open]').forEach(function(d){
      if(!d.contains(e.target)||e.target.closest('.ahqs-mobile-links a'))d.removeAttribute('open');
    });
  });
  // Back-to-top button for long pages
  var b=document.createElement('button');b.type='button';b.className='ahqs-to-top';b.setAttribute('aria-label','Back to top of page');b.textContent='\u2191';
  b.addEventListener('click',function(){window.scrollTo({top:0,behavior:'smooth'})});
  document.addEventListener('DOMContentLoaded',function(){document.body.appendChild(b)});
  if(document.body)document.body.appendChild(b);
  window.addEventListener('scroll',function(){b.classList.toggle('show',window.scrollY>900)},{passive:true});
})();
