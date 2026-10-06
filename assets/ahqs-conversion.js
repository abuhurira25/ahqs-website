/* AHQS Conversion Enhancement System — 2026
   Handles: lead capture form submission, assessment tool scoring,
   newsletter signup, social share, scroll-depth tracking.
   NOTE: This is a static GitHub Pages site — form submissions construct
   mailto: links. For production, connect Formspree, Getform, or a CRM API. */

(function() {
  'use strict';

  // ===== GA4 EVENT HELPER =====
  function send(name, params) {
    try { if (typeof gtag === 'function') gtag('event', name, params || {}); } catch(e) {}
  }

  // ===== LEAD CAPTURE FORM HANDLING =====
  document.querySelectorAll('.ahqs-lead-form').forEach(function(form) {
    form.addEventListener('submit', function(e) {
      e.preventDefault();
      var name = (form.querySelector('[name="name"]') || {}).value || '';
      var email = (form.querySelector('[name="email"]') || {}).value || '';
      var orgType = (form.querySelector('[name="org_type"]') || {}).value || '';
      var resource = form.getAttribute('data-resource') || 'AHQS Resource';
      var consent = form.querySelector('[name="consent"]');
      if (consent && !consent.checked) {
        alert('Please agree to receive emails from AHQS before submitting.');
        return;
      }
      if (!email) { alert('Please enter your email address.'); return; }

      // Construct mailto link for static site
      var subject = encodeURIComponent('Resource Request: ' + resource);
      var body = encodeURIComponent(
        'Name: ' + name + '\n' +
        'Email: ' + email + '\n' +
        'Organization Type: ' + orgType + '\n' +
        'Resource Requested: ' + resource + '\n\n' +
        'Please send the requested resource. I consent to receiving emails from AHQS.'
      );
      var mailto = 'mailto:info@ahqshealthcare.com?subject=' + subject + '&body=' + body;

      // Show success message
      var successMsg = form.parentElement.querySelector('.ahqs-lead-success');
      if (successMsg) {
        successMsg.classList.add('ahqs-show');
        successMsg.querySelector('span').textContent = email;
      }

      // Track event
      send('lead_form_submitted', {
        resource_name: resource,
        org_type: orgType,
        link_text: name
      });

      // Open email client
      window.location.href = mailto;
    });
  });

  // ===== NEWSLETTER SIGNUP =====
  document.querySelectorAll('.ahqs-newsletter-form').forEach(function(form) {
    form.addEventListener('submit', function(e) {
      e.preventDefault();
      var email = (form.querySelector('[name="email"]') || {}).value || '';
      if (!email) { alert('Please enter your email address.'); return; }

      var subject = encodeURIComponent('AHQS Newsletter Signup');
      var body = encodeURIComponent('Please add me to the AHQS newsletter list.\n\nEmail: ' + email);
      var mailto = 'mailto:info@ahqshealthcare.com?subject=' + subject + '&body=' + body;

      send('newsletter_signup', { link_text: email });

      var btn = form.querySelector('button');
      if (btn) { btn.textContent = 'Subscribed!'; btn.disabled = true; }
      var input = form.querySelector('input');
      if (input) { input.value = ''; input.placeholder = 'You are subscribed!'; }

      window.location.href = mailto;
    });
  });

  // ===== INTERACTIVE ASSESSMENT TOOL =====
  var assessments = document.querySelectorAll('.ahqs-assessment');
  assessments.forEach(function(assessment) {
    var typeBtns = assessment.querySelectorAll('.ahqs-assessment-type-btn');
    var questionsSection = assessment.querySelector('.ahqs-assessment-questions');
    var questions = assessment.querySelectorAll('.ahqs-assessment-q');
    var resultSection = assessment.querySelector('.ahqs-assessment-result');
    var progressBar = assessment.querySelector('.ahqs-assessment-progress-bar');
    var scoreDisplay = assessment.querySelector('.ahqs-assessment-score');
    var scoreLabel = assessment.querySelector('.ahqs-assessment-score-label');
    var gapsList = assessment.querySelector('.ahqs-assessment-gaps ul');
    var currentQ = 0;
    var answers = [];
    var selectedType = '';

    // Assessment questions data
    var questionBank = {
      iso15189: [
        { q: 'Is your Quality Management System (QMS) fully documented and implemented?', area: 'Quality Management System' },
        { q: 'Are personnel competency records current and verifiable?', area: 'Personnel Competency' },
        { q: 'Is document control consistent across all departments?', area: 'Document Control' },
        { q: 'Are equipment calibration and metrological traceability records complete?', area: 'Equipment & Metrology' },
        { q: 'Are pre-examination, examination, and post-examination processes validated?', area: 'Examination Processes' },
        { q: 'Is risk management integrated into daily laboratory practice?', area: 'Risk Management' },
        { q: 'Are internal audits conducted regularly with documented findings?', area: 'Internal Audit' },
        { q: 'Are CAPA records effective and sustainable (no recurring findings)?', area: 'CAPA & Improvement' },
        { q: 'Is management review conducted and documented annually?', area: 'Management Review' },
        { q: 'Have you conducted a mock assessment in the last 12 months?', area: 'Assessment Readiness' }
      ],
      cap: [
        { q: 'Are all CAP checklist requirements assigned to specific owners?', area: 'Checklist Ownership' },
        { q: 'Is evidence readily retrievable for all inspector questions?', area: 'Evidence Readiness' },
        { q: 'Are competency assessments documented for all testing personnel?', area: 'Competency' },
        { q: 'Are QC and PT records complete with timely review?', area: 'QC/PT Oversight' },
        { q: 'Are corrective actions tracked to closure with effectiveness checks?', area: 'Corrective Actions' },
        { q: 'Have you conducted a mock inspection in the last 12 months?', area: 'Mock Inspection' },
        { q: 'Is your procedure manual current with all CAP changes?', area: 'Procedure Manual' },
        { q: 'Are instrument records and maintenance logs complete?', area: 'Instrument Records' }
      ],
      jci: [
        { q: 'Is leadership governance structure documented and aligned with JCI standards?', area: 'Leadership & Governance' },
        { q: 'Are patient safety goals integrated into daily practice?', area: 'Patient Safety' },
        { q: 'Are clinical audit and tracer methodologies in use?', area: 'Clinical Audit' },
        { q: 'Is risk management integrated across all clinical areas?', area: 'Risk Management' },
        { q: 'Are credentialing and privileging processes current?', area: 'Credentialing' },
        { q: 'Is evidence organized and retrieval-ready for surveyors?', area: 'Evidence Readiness' },
        { q: 'Have you conducted a mock survey in the last 12 months?', area: 'Mock Survey' },
        { q: 'Are quality indicators monitored and acted upon?', area: 'Quality Indicators' }
      ],
      aabb: [
        { q: 'Is your blood bank quality system fully documented?', area: 'Quality System' },
        { q: 'Are competency records current for all transfusion service staff?', area: 'Competency' },
        { q: 'Is traceability documented for all blood products?', area: 'Traceability' },
        { q: 'Are process controls validated and documented?', area: 'Process Control' },
        { q: 'Are internal audits conducted and findings tracked?', area: 'Internal Audit' },
        { q: 'Are corrective actions effective with no recurrence?', area: 'Corrective Actions' },
        { q: 'Is risk assessment integrated into transfusion practices?', area: 'Risk Management' },
        { q: 'Have you prepared for survey readiness in the last 12 months?', area: 'Survey Readiness' }
      ]
    };

    typeBtns.forEach(function(btn) {
      btn.addEventListener('click', function() {
        typeBtns.forEach(function(b) { b.setAttribute('aria-selected', 'false'); });
        btn.setAttribute('aria-selected', 'true');
        selectedType = btn.dataset.type;

        // Load questions
        var qs = questionBank[selectedType] || [];
        questionsSection.innerHTML = '';
        answers = new Array(qs.length).fill(null);
        currentQ = 0;

        qs.forEach(function(q, i) {
          var div = document.createElement('div');
          div.className = 'ahqs-assessment-q' + (i === 0 ? ' ahqs-active' : '');
          div.innerHTML = '<h4>' + (i + 1) + '. ' + q.q + '</h4>' +
            '<div class="ahqs-assessment-options">' +
              '<label class="ahqs-assessment-opt"><input type="radio" name="q' + i + '" value="3"><span>Fully implemented and evidenced</span></label>' +
              '<label class="ahqs-assessment-opt"><input type="radio" name="q' + i + '" value="2"><span>Partially implemented</span></label>' +
              '<label class="ahqs-assessment-opt"><input type="radio" name="q' + i + '" value="1"><span>Not implemented or unsure</span></label>' +
            '</div>';
          questionsSection.appendChild(div);
        });

        // Add navigation
        var nav = document.createElement('div');
        nav.className = 'ahqs-assessment-nav';
        nav.innerHTML = '<button class="btn btn-outline ahqs-prev-btn" type="button" disabled>Previous</button>' +
          '<span class="ahqs-q-counter"></span>' +
          '<button class="btn btn-gold ahqs-next-btn" type="button">Next →</button>';
        questionsSection.appendChild(nav);

        questions = questionsSection.querySelectorAll('.ahqs-assessment-q');
        updateProgress();

        questionsSection.classList.add('ahqs-show');

        // Nav handlers
        var prevBtn = questionsSection.querySelector('.ahqs-prev-btn');
        var nextBtn = questionsSection.querySelector('.ahqs-next-btn');
        var counter = questionsSection.querySelector('.ahqs-q-counter');

        prevBtn.addEventListener('click', function() {
          if (currentQ > 0) {
            questions[currentQ].classList.remove('ahqs-active');
            currentQ--;
            questions[currentQ].classList.add('ahqs-active');
            updateProgress();
          }
        });

        nextBtn.addEventListener('click', function() {
          // Check if current question is answered
          var checked = questions[currentQ].querySelector('input[type="radio"]:checked');
          if (!checked) {
            alert('Please select an answer for this question.');
            return;
          }
          answers[currentQ] = parseInt(checked.value);

          if (currentQ < questions.length - 1) {
            questions[currentQ].classList.remove('ahqs-active');
            currentQ++;
            questions[currentQ].classList.add('ahqs-active');
            updateProgress();
          } else {
            showResults();
          }
        });

        function updateProgress() {
          var pct = ((currentQ + 1) / questions.length) * 100;
          if (progressBar) progressBar.style.width = pct + '%';
          if (counter) counter.textContent = (currentQ + 1) + ' of ' + questions.length;
          if (prevBtn) prevBtn.disabled = (currentQ === 0);
          if (nextBtn) nextBtn.textContent = (currentQ === questions.length - 1) ? 'See Results →' : 'Next →';
        }

        function showResults() {
          var total = answers.reduce(function(a, b) { return a + (b || 0); }, 0);
          var maxScore = questions.length * 3;
          var pct = Math.round((total / maxScore) * 100);
          var qs = questionBank[selectedType] || [];
          var gaps = [];
          answers.forEach(function(val, i) {
            if (val < 3) gaps.push(qs[i].area);
          });

          if (scoreDisplay) scoreDisplay.textContent = pct + '%';
          if (scoreLabel) {
            if (pct >= 80) scoreLabel.textContent = 'Strong readiness — minor gaps to address';
            else if (pct >= 50) scoreLabel.textContent = 'Moderate readiness — several gaps need attention';
            else scoreLabel.textContent = 'Early readiness — significant gaps identified';
          }
          if (gapsList) {
            gapsList.innerHTML = gaps.length > 0
              ? gaps.map(function(g) { return '<li>' + g + '</li>'; }).join('')
              : '<li style="border:0;color:#087f83">No significant gaps identified — well done!</li>';
          }

          questionsSection.style.display = 'none';
          resultSection.classList.add('ahqs-show');

          send('readiness_check_completed', {
            accreditation_type: selectedType,
            score: pct,
            gaps_count: gaps.length
          });
        }
      });
    });
  });

  // ===== DOWNLOAD BUTTON TRACKING =====
  document.querySelectorAll('.ahqs-download-btn').forEach(function(btn) {
    btn.addEventListener('click', function(e) {
      var resource = btn.getAttribute('data-resource') || 'Unknown Resource';
      send('guide_download_requested', { resource_name: resource });
    });
  });

  // ===== SOCIAL SHARE =====
  document.querySelectorAll('.ahqs-share-btn').forEach(function(btn) {
    btn.addEventListener('click', function(e) {
      e.preventDefault();
      var platform = btn.dataset.platform;
      var url = encodeURIComponent(window.location.href);
      var title = encodeURIComponent(document.title);
      var shareUrl = '';
      if (platform === 'linkedin') {
        shareUrl = 'https://www.linkedin.com/sharing/share-offsite/?url=' + url;
      } else if (platform === 'twitter') {
        shareUrl = 'https://twitter.com/intent/tweet?url=' + url + '&text=' + title;
      } else if (platform === 'email') {
        shareUrl = 'mailto:?subject=' + title + '&body=I thought you might find this useful: ' + decodeURIComponent(url);
      } else if (platform === 'copy') {
        navigator.clipboard.writeText(window.location.href).then(function() {
          btn.textContent = '✓';
          setTimeout(function() { btn.textContent = '🔗'; }, 2000);
        });
        return;
      }
      if (shareUrl) window.open(shareUrl, '_blank', 'noopener,noreferrer,width=600,height=500');
      send('social_share_click', { platform: platform });
    });
  });

  // ===== SCROLL DEPTH TRACKING (50% and 100%) =====
  var scroll50Sent = false, scroll100Sent = false;
  if (document.querySelector('.article')) {
    document.addEventListener('scroll', function() {
      var d = document.documentElement;
      var max = d.scrollHeight - d.clientHeight;
      var pct = max > 0 ? (d.scrollTop / max) * 100 : 0;
      if (pct >= 50 && !scroll50Sent) { scroll50Sent = true; send('guide_read_50', {}); }
      if (pct >= 95 && !scroll100Sent) { scroll100Sent = true; send('guide_read_100', {}); }
    }, { passive: true });
  }
})();
