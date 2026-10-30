(function () {
  'use strict';

  // モバイルナビの開閉
  var toggle = document.getElementById('navToggle');
  var nav = document.getElementById('navMenu');

  if (toggle && nav) {
    toggle.addEventListener('click', function () {
      var isOpen = nav.classList.toggle('is-open');
      toggle.setAttribute('aria-expanded', String(isOpen));
      toggle.setAttribute('aria-label', isOpen ? 'メニューを閉じる' : 'メニューを開く');
    });

    // リンクをタップしたらメニューを閉じる(モバイル)
    nav.querySelectorAll('a').forEach(function (link) {
      link.addEventListener('click', function () {
        nav.classList.remove('is-open');
        toggle.setAttribute('aria-expanded', 'false');
        toggle.setAttribute('aria-label', 'メニューを開く');
      });
    });
  }

  // ---------- スポンサー一覧(Excel由来データ)の描画 + 詳細モーダル ----------
  var sectionsContainer = document.getElementById('sponsorsSections');
  var modal = document.getElementById('sponsorModal');
  if (!sectionsContainer || !modal || typeof window.loadSponsors !== 'function') { return; }

  var GROUP_ORDER = ['topteam', 'kobekatsu'];
  var GROUP_FALLBACK_LABEL = { topteam: 'トップチームスポンサー', kobekatsu: 'コベカツクラブ' };
  var TIER_ORDER = ['main', 'gold', 'silver', 'bronze', 'supplier', 'partner'];
  var TIER_CLASS = {
    main: 'tier--gold', gold: 'tier--gold', silver: 'tier--silver',
    bronze: 'tier--bronze', supplier: 'tier--supplier'
  };

  function groupBy(list, keyFn) {
    var map = {};
    var order = [];
    list.forEach(function (item) {
      var key = keyFn(item);
      if (!map[key]) { map[key] = []; order.push(key); }
      map[key].push(item);
    });
    return { map: map, order: order };
  }

  function sortByPriority(keys, priority) {
    return keys.slice().sort(function (a, b) {
      var ia = priority.indexOf(a);
      var ib = priority.indexOf(b);
      if (ia === -1) { ia = priority.length; }
      if (ib === -1) { ib = priority.length; }
      if (ia !== ib) { return ia - ib; }
      return a < b ? -1 : (a > b ? 1 : 0);
    });
  }

  function el(tag, className, html) {
    var node = document.createElement(tag);
    if (className) { node.className = className; }
    if (html !== undefined) { node.innerHTML = html; }
    return node;
  }

  function escapeHtml(str) {
    return String(str || '').replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function buildSponsorCard(sponsor) {
    var card = el('article', 'sponsor-card');
    card.dataset.sponsorId = sponsor.id;

    var logo = el('div', 'sponsor-card__logo');
    if (sponsor.images.length) {
      var img = document.createElement('img');
      img.src = sponsor.images[0];
      img.alt = '';
      img.loading = 'lazy';
      logo.appendChild(img);
    } else {
      logo.setAttribute('aria-hidden', 'true');
      logo.textContent = sponsor.name;
    }
    card.appendChild(logo);

    card.appendChild(el('h4', 'sponsor-card__name', escapeHtml(sponsor.name)));
    card.appendChild(el('p', 'sponsor-card__benefit', escapeHtml(sponsor.benefit || '詳細はこちら')));

    var more = document.createElement('button');
    more.type = 'button';
    more.className = 'sponsor-card__more';
    more.innerHTML = '詳細を見る <span aria-hidden="true">&rsaquo;</span>';
    card.appendChild(more);

    return card;
  }

  function renderSponsorSections(sponsors) {
    sectionsContainer.innerHTML = '';
    var byGroup = groupBy(sponsors, function (s) { return s.group; });
    var groupKeys = sortByPriority(byGroup.order, GROUP_ORDER);

    groupKeys.forEach(function (groupKey) {
      var groupSponsors = byGroup.map[groupKey];
      var groupLabel = (groupSponsors[0] && groupSponsors[0].groupLabel) || GROUP_FALLBACK_LABEL[groupKey] || groupKey;

      var groupSection = el('div', 'sponsor-group');
      groupSection.appendChild(el('h3', 'sponsor-group__title', escapeHtml(groupLabel)));

      var byTier = groupBy(groupSponsors, function (s) { return s.tier; });
      var hasTierLabels = groupSponsors.some(function (s) { return s.tierLabel; });

      if (hasTierLabels) {
        var tierKeys = sortByPriority(byTier.order, TIER_ORDER);
        tierKeys.forEach(function (tierKey) {
          var tierSponsors = byTier.map[tierKey];
          var tierLabel = (tierSponsors[0] && tierSponsors[0].tierLabel) || '';
          var tierWrap = el('div', 'tier ' + (TIER_CLASS[tierKey] || ''));
          if (tierLabel) {
            var tierTitle = el('h4', 'tier__title');
            tierTitle.appendChild(el('span', 'tier__badge', escapeHtml(tierLabel)));
            tierWrap.appendChild(tierTitle);
          }
          var grid = el('div', 'sponsor-grid');
          tierSponsors.forEach(function (s) { grid.appendChild(buildSponsorCard(s)); });
          tierWrap.appendChild(grid);
          groupSection.appendChild(tierWrap);
        });
      } else {
        var flatGrid = el('div', 'sponsor-grid');
        groupSponsors.forEach(function (s) { flatGrid.appendChild(buildSponsorCard(s)); });
        groupSection.appendChild(flatGrid);
      }

      sectionsContainer.appendChild(groupSection);
    });
  }

  // ---------- 詳細モーダル ----------
  var els = {
    galleryTrack: document.getElementById('sponsorModalGalleryTrack'),
    galleryPrev: document.getElementById('galleryPrev'),
    galleryNext: document.getElementById('galleryNext'),
    galleryDots: document.getElementById('sponsorModalGalleryDots'),
    galleryStatus: document.getElementById('sponsorModalGalleryStatus'),
    tier: document.getElementById('sponsorModalTier'),
    name: document.getElementById('sponsorModalName'),
    address: document.getElementById('sponsorModalAddress'),
    desc: document.getElementById('sponsorModalDesc'),
    benefit: document.getElementById('sponsorModalBenefit'),
    links: document.getElementById('sponsorModalLinks')
  };
  var closeBtn = document.getElementById('sponsorModalClose');
  var lastFocused = null;
  var galleryObserver = null;
  var galleryCount = 0;

  var linkDefs = [
    { key: 'website', label: null },
    { key: 'instagram', label: 'Instagram' },
    { key: 'line', label: '公式LINE' },
    { key: 'mapUrl', label: '地図を見る' }
  ];

  function fillText(target, text, fallback) {
    if (text) {
      target.textContent = text;
      target.hidden = false;
    } else if (fallback) {
      target.textContent = fallback;
      target.hidden = false;
    } else {
      target.hidden = true;
    }
  }

  function setActiveSlide(index) {
    els.galleryDots.querySelectorAll('.sponsor-modal__gallery-dot').forEach(function (dot) {
      dot.classList.toggle('is-active', Number(dot.dataset.index) === index);
    });
    updateGalleryStatus(index, galleryCount);
  }

  function updateGalleryStatus(index, total) {
    els.galleryStatus.textContent = total > 1 ? (index + 1) + ' / ' + total + '枚目' : '';
  }

  function goToSlide(index) {
    var slide = els.galleryTrack.children[index];
    if (slide) {
      slide.scrollIntoView({ behavior: 'smooth', inline: 'start', block: 'nearest' });
    }
  }

  function shiftSlide(dir) {
    var active = els.galleryDots.querySelector('.sponsor-modal__gallery-dot.is-active');
    var current = active ? Number(active.dataset.index) : 0;
    var next = (current + dir + galleryCount) % galleryCount;
    goToSlide(next);
  }

  function buildGallery(images, name) {
    if (galleryObserver) {
      galleryObserver.disconnect();
      galleryObserver = null;
    }
    els.galleryTrack.innerHTML = '';
    els.galleryDots.innerHTML = '';

    var list = (images && images.length) ? images : [null];
    galleryCount = list.length;
    var multi = list.length > 1;

    list.forEach(function (src, index) {
      var slide = document.createElement('div');
      slide.className = 'sponsor-modal__gallery-slide';
      slide.setAttribute('role', 'group');
      slide.setAttribute('aria-roledescription', 'スライド');
      slide.setAttribute('aria-label', (index + 1) + ' / ' + list.length);
      slide.dataset.index = String(index);

      if (src) {
        var img = document.createElement('img');
        img.src = src;
        img.alt = name + 'の画像 ' + (index + 1);
        img.loading = index === 0 ? 'eager' : 'lazy';
        slide.appendChild(img);
      } else {
        slide.textContent = '画像は準備中です';
      }

      if (multi) {
        slide.classList.add('is-clickable');
        slide.addEventListener('click', function () { shiftSlide(1); });
      }

      els.galleryTrack.appendChild(slide);

      if (multi) {
        var dot = document.createElement('button');
        dot.type = 'button';
        dot.className = 'sponsor-modal__gallery-dot';
        dot.setAttribute('aria-label', (index + 1) + '枚目を表示');
        dot.dataset.index = String(index);
        if (index === 0) { dot.classList.add('is-active'); }
        dot.addEventListener('click', function (e) {
          e.stopPropagation();
          goToSlide(index);
        });
        els.galleryDots.appendChild(dot);
      }
    });

    els.galleryPrev.hidden = !multi;
    els.galleryNext.hidden = !multi;
    els.galleryDots.hidden = !multi;
    updateGalleryStatus(0, list.length);

    if (multi) {
      galleryObserver = new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting && entry.intersectionRatio >= 0.6) {
            setActiveSlide(Number(entry.target.dataset.index));
          }
        });
      }, { root: els.galleryTrack, threshold: 0.6 });

      els.galleryTrack.querySelectorAll('.sponsor-modal__gallery-slide').forEach(function (slide) {
        galleryObserver.observe(slide);
      });
    }
  }

  function openSponsorModal(data) {
    buildGallery(data.images, data.name);

    els.tier.textContent = data.tierLabel || data.groupLabel || '';
    els.name.textContent = data.name;
    fillText(els.address, data.address ? '所在地: ' + data.address : null, null);
    fillText(els.desc, data.description, null);
    fillText(els.benefit, data.benefit, null);

    els.links.innerHTML = '';
    var hasLink = false;
    linkDefs.forEach(function (def) {
      var url = data[def.key];
      if (!url) { return; }
      hasLink = true;
      var a = document.createElement('a');
      a.href = url;
      a.target = '_blank';
      a.rel = 'noopener';
      a.className = 'sponsor-modal__link';
      a.textContent = def.label || data.websiteLabel || '公式サイト';
      els.links.appendChild(a);
    });
    if (!hasLink) {
      els.links.appendChild(el('p', 'sponsor-modal__links-empty', 'リンク情報は準備中です'));
    }

    lastFocused = document.activeElement;
    modal.showModal();
    document.body.classList.add('modal-open');
    closeBtn.focus();
  }

  function closeSponsorModal() {
    modal.close();
  }

  closeBtn.addEventListener('click', closeSponsorModal);
  els.galleryPrev.addEventListener('click', function () { shiftSlide(-1); });
  els.galleryNext.addEventListener('click', function () { shiftSlide(1); });

  modal.addEventListener('click', function (e) {
    if (e.target === modal) { closeSponsorModal(); }
  });

  modal.addEventListener('keydown', function (e) {
    if (galleryCount <= 1) { return; }
    if (e.key === 'ArrowLeft') { shiftSlide(-1); }
    if (e.key === 'ArrowRight') { shiftSlide(1); }
  });

  modal.addEventListener('close', function () {
    document.body.classList.remove('modal-open');
    if (galleryObserver) {
      galleryObserver.disconnect();
      galleryObserver = null;
    }
    if (lastFocused && typeof lastFocused.focus === 'function') {
      lastFocused.focus();
    }
  });

  window.loadSponsors().then(function (result) {
    var sponsors = result.sponsors;
    renderSponsorSections(sponsors);

    sectionsContainer.querySelectorAll('.sponsor-card').forEach(function (card) {
      card.addEventListener('click', function () {
        var data = sponsors.find(function (s) { return s.id === card.dataset.sponsorId; });
        if (data) { openSponsorModal(data); }
      });
    });

    if (result.source === 'fallback') {
      var notice = el('p', 'sponsors-notice', '※ 最新データの取得に失敗したため、保存されている内容を表示しています。');
      sectionsContainer.parentNode.insertBefore(notice, sectionsContainer);
    }
  });
})();
