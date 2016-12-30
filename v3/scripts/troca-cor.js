function Trocar() {
  // define a cor da página
  cor = Math.floor(Math.random() * 4) + 1;

  // altera a cor do fundo da página
  switch (cor) {
    case 1:
      document.body.style.backgroundColor = '#FFFFD0'; break;
    case 2:
      document.body.style.backgroundColor = '#D0FFDE'; break;
    case 3:
      document.body.style.backgroundColor = '#CCFFFF'; break;
    case 4:
      document.body.style.backgroundColor = '#FFCCCC'; break;
  }

  // define a logo de acordo com a cor
  document.getElementById('logo').src = 'gifs/logo' + cor + '.gif';

  // escolhe o banner a ser mostrado no topo da página
  var todosBanners = [
    'linhazinhas',
    'sabeoqe',
    'saladois',
    'especiais'
  ];
  var bannerEscolhido = todosBanners[Math.floor(Math.random() * todosBanners.length)];

  // document.macflash.src = 'flashes/' + bannerEscolhido + cor + '.swf';

  var flashEl = document.getElementById('banner-top-flash');
  var videoEl = document.getElementById('banner-top-video');

  flashEl.getElementsByTagName('param')[0].value = 'flashes\\' + bannerEscolhido + cor + '.swf';

  ['webm', 'mp4'].forEach(function(type) {
    var sourceEl = document.createElement('source');
    sourceEl.setAttribute('type', 'video/' + type);
    sourceEl.setAttribute('src', 'flashes/' + bannerEscolhido + cor + '.' + type);

    videoEl.appendChild(sourceEl);
  });
}
