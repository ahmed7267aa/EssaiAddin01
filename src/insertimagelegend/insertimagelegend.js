// Ce code est CRUCIAL - il attend que Office.js soit prêt
Office.initialize = function() {
  console.log("Office.js est initialisé");
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Word est prêt ! Votre complément peut s'initialiser");

    // Fonction utilitaire de log (écrit en HTML dans les zones identifiées)
    function log(id, txt, level) {
      const el = document.getElementById(id);
      if (!el) {
        console.error(`Élément avec ID "${id}" non trouvé`);
        return;
      }

      const span = document.createElement('div');
      span.innerHTML = txt;
      el.prepend(span);

      const consoleEl = document.getElementById('console');
      if (consoleEl) {
        consoleEl.textContent = txt + '\n' + consoleEl.textContent;
      }
    }

    // Core: get OOXML and parse
    async function getBodyOoxml() {
      return Word.run(async (context) => {
        const body = context.document.body;
        const ooxml = body.getOoxml();
        await context.sync();
        return ooxml.value;
      });
    }

    // Helper: extract text content from an OOXML <w:p> element
    function getParagraphText(pNode) {
      // concatenate all w:t child text nodes
      const texts = pNode.getElementsByTagName('w:t');
      let s = '';
      for (let i = 0; i < texts.length; i++) s += texts[i].textContent;
      return s.trim();
    }

    // Parse OOXML and find images & captions (version améliorée pour la "liste des figures")
    async function analyzeDocument() {
      try {
        const xmlString = await getBodyOoxml();
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlString, 'application/xml');

        // Vérifier si le parsing a réussi
        const parserError = doc.getElementsByTagName('parsererror');
        if (parserError.length > 0) {
          throw new Error('Erreur lors de l\'analyse OOXML: ' + parserError[0].textContent);
        }

        // Count images: look for <w:drawing> or <w:pict>
        const drawings = doc.getElementsByTagName('w:drawing') || [];
        const picts = doc.getElementsByTagName('w:pict') || [];
        const imageCount = (drawings.length || 0) + (picts.length || 0);

        // Find all paragraph nodes
        const paragraphs = doc.getElementsByTagName('w:p');

        // Find caption paragraphs: text beginning with 'Figure <num>' (French)
        const captionRegex = /^Figure\s+\d+\s*[:\-–]?/i;
        const captions = [];
        for (let i = 0; i < paragraphs.length; i++) {
          const p = paragraphs[i];
          const txt = getParagraphText(p);
          if (captionRegex.test(txt)) {
            // check for centering: look for <w:jc w:val="center"/> inside <w:pPr>
            let centered = false;
            const pPr = p.getElementsByTagName('w:pPr');
            if (pPr && pPr.length > 0) {
              const jc = pPr[0].getElementsByTagName('w:jc');
              if (jc && jc.length > 0 && jc[0].getAttribute('w:val') && jc[0].getAttribute('w:val').toLowerCase() === 'center') centered = true;
            }
            captions.push({ text: txt, centered: centered, nodeIndex: i });
          }
        }

        // --------------------------------------------------------------------
        // Find "Liste des figures" (heading) and its following entries — version plus robuste
        // --------------------------------------------------------------------
        let listNodeIndex = -1;
        let listEntries = [];

        // Helper to find paragraph index from a node (walk up to w:p)
        function findParagraphIndexFromNode(node, paragraphs) {
          let anc = node;
          while (anc && anc.nodeName && anc.nodeName.toLowerCase() !== 'w:p') anc = anc.parentNode;
          if (!anc) return -1;
          for (let i = 0; i < paragraphs.length; i++) {
            if (paragraphs[i] === anc) return i;
          }
          return -1;
        }

        // 1) Rechercher des champs/field instructions (TOC / Table of Figures) dans le OOXML
        const instrNodes = Array.from(doc.getElementsByTagName('w:instrText') || []);
        for (let k = 0; k < instrNodes.length && listNodeIndex === -1; k++) {
          const txt = (instrNodes[k].textContent || '').trim();
          if (/TOC\b|tableof|table of figures|table des illustrations|liste des figures|table des figures/i.test(txt)) {
            listNodeIndex = findParagraphIndexFromNode(instrNodes[k], paragraphs);
            break;
          }
        }

        // 2) Rechercher des <w:fldSimple w:instr="..."> si pas trouvé
        if (listNodeIndex === -1) {
          const fldSimples = Array.from(doc.getElementsByTagName('w:fldSimple') || []);
          for (let k = 0; k < fldSimples.length; k++) {
            const instrAttr = (fldSimples[k].getAttribute && (fldSimples[k].getAttribute('w:instr') || fldSimples[k].getAttribute('instr'))) || '';
            if (/TOC\b|tableof|table of figures|table des illustrations|liste des figures|table des figures/i.test(instrAttr)) {
              listNodeIndex = findParagraphIndexFromNode(fldSimples[k], paragraphs);
              break;
            }
          }
        }

        // 3) Si toujours pas trouvé, heuristique : détecter un paragraphe qui "ressemble" à une entrée de table des figures
        if (listNodeIndex === -1) {
          for (let i = 0; i < paragraphs.length; i++) {
            const txt = getParagraphText(paragraphs[i]);
            if (!txt) continue;
            // candidate: "Figure 1 : Titre .... 2" ou "Figure 1\tTitre\t2"
            if (/\bFigure\s*\d+\b/i.test(txt) && (/\.{3,}/.test(txt) || /\s\d+$/.test(txt) || txt.indexOf('\t') !== -1)) {
              listNodeIndex = i;
              break;
            }
          }
        }

        // 4) Si on a trouvé une position, collecter les entrées suivantes (jusqu'à paragraphe vide ou limite)
        if (listNodeIndex !== -1) {
          for (let j = listNodeIndex + 1; j < paragraphs.length; j++) {
            const txt = getParagraphText(paragraphs[j]);
            if (!txt) break; // arrêt sur paragraphe vide (heuristique)
            // considérer comme entrée si ressemble à "Figure ..." OU contient 'figure' et un numéro/page
            if (captionRegex.test(txt) || (txt.toLowerCase().includes('figure') && (/\s\d+$/.test(txt) || /\.{3,}/.test(txt) || txt.indexOf('\t') !== -1))) {
              const hyperlinks = paragraphs[j].getElementsByTagName('w:hyperlink');
              const hasLink = hyperlinks && hyperlinks.length > 0;
              listEntries.push({ text: txt, hasLink: hasLink });
            } else {
              // si paragraphe court / pas ressemblant on arrête (heuristique)
              if (txt.length < 2) break;
              // sinon on continue mais ne l'ajoute pas
            }
            // safety: ne pas scanner indéfiniment
            if (listEntries.length > 200) break;
          }
        }

        // Retourner le résultat
        return { imageCount, captions, listNodeIndex, listEntries };
      } catch (err) {
        console.error("Erreur dans analyzeDocument:", err);
        throw err;
      }
    }

    // --------------------------------------------------------------------
    // Écouteurs des boutons (assume que les éléments existent dans le DOM)
    // --------------------------------------------------------------------
    document.addEventListener('DOMContentLoaded', () => {
      // bouton - images
      const chkImages = document.getElementById('chk-images');
      if (chkImages) {
        chkImages.onclick = async function() {
          try {
            const res = await analyzeDocument();

            if (res.imageCount < 2) {
              if (res.imageCount === 0) {
                log('res-images', '<span class="ko">Aucune image détectée dans le document.</span>');
                log('console', 'Vérification images: 0 trouvées.');
              } else {
                log('res-images', '<span class="ko">Pas plusieurs images détectées (seulement 1 image).</span>');
                log('console', 'Vérification images: 1 trouvée — pas de pluralité.');
              }
            } else if (res.imageCount === 2) {
              // Message en vert pour indiquer que c'est bon quand il y a exactement 2 images
              log('res-images', '<span class="ok">Images détectées: 2 — OK.</span>');
              log('console', 'Vérification images: 2 trouvées — OK.');
            } else {
              // Plus de 2 images
              log('res-images', 'Images détectées: ' + res.imageCount);
              log('console', 'Vérification images: ' + res.imageCount + ' trouvées.');
            }
          } catch (err) {
            log('res-images', 'Erreur: ' + err.message);
            console.error(err);
          }
        };
      }

      // bouton - légendes
      const chkCaptions = document.getElementById('chk-captions');
      if (chkCaptions) {
        chkCaptions.onclick = async function() {
          try {
            const res = await analyzeDocument();
            const capCount = res.captions.length;
            if (capCount === 0) {
              log('res-captions', 'Aucune légende détectée (pattern "Figure N : ...").');
            } else {
              log('res-captions', 'Légendes détectées: ' + capCount);
              res.captions.forEach((c, i) => {
                const cent = c.centered ? '<span class="ok">centrée</span>' : '<span class="ko">non centrée</span>';
                log('res-captions', '• ' + c.text + ' — ' + cent);
              });
            }
            log('console', 'Vérification légendes terminée.');
          } catch (err) {
            log('res-captions', 'Erreur: ' + err.message);
            console.error(err);
          }
        };
      }

      // bouton - liste des figures
// bouton - liste des figures (remplace l'ancien chkList.onclick par ce bloc)
const chkList = document.getElementById('chk-list');
if (chkList) {
  chkList.onclick = async function() {
    try {
      const res = await analyzeDocument();

      // Pas de table détectée
      if (res.listNodeIndex === -1) {
        log('res-list', '<span class="ko">Aucune "Liste des figures" détectée.</span>');
        log('console', 'Vérification liste des figures: introuvable.');
        return;
      }

      // Table trouvée — analyser les entrées
      if (!res.listEntries || res.listEntries.length === 0) {
        log('res-list', '<span class="ko">Table des figures trouvée mais aucune entrée détectée.</span>');
        log('console', 'Vérification liste des figures: table trouvée mais 0 entrées.');
        return;
      }

      // Condition "respectée" : nombre d'entrées == nombre de légendes
      const countsMatch = (res.listEntries.length === res.captions.length);
      if (countsMatch) {
        log('res-list', '<span class="ok">Liste des figures trouvée et nombre d\'entrées cohérent avec les légendes (' + res.captions.length + ').</span>');
        log('console', 'Vérification liste des figures: OK (' + res.listEntries.length + ' entrées).');
      } else {
        log('res-list', '<span class="ko">Liste trouvée mais nombre d\'entrées (' + res.listEntries.length + ") != nombre de légendes ('" + res.captions.length + "').</span>");
        log('console', 'Vérification liste des figures: KO (comptes différents).');
      }

      // Afficher les entrées — si tout est OK, on les met en vert ; sinon on colore selon le lien (ou rouge)
      res.listEntries.forEach((e, i) => {
        const lineClass = countsMatch ? 'ok' : (e.hasLink ? 'ok' : 'ko');
        log('res-list', '• <span class="' + lineClass + '">' + e.text + '</span>' + (e.hasLink ? ' (lien)' : ' (pas de lien)'));
      });

    } catch (err) {
      log('res-list', 'Erreur: ' + err.message);
      console.error(err);
    }
  };
}


      // bouton - liens dans la liste
      const chkLinks = document.getElementById('chk-links');
      if (chkLinks) {
        chkLinks.onclick = async function() {
          try {
            const res = await analyzeDocument();
            if (res.listNodeIndex === -1) {
              log('res-links', 'Impossible de vérifier les liens car la liste des figures est introuvable.');
              return;
            }
            const missing = res.listEntries.filter(e => !e.hasLink);
            if (missing.length === 0) {
              log('res-links', 'Tous les éléments de la liste des figures semblent avoir des liens.');
            } else {
              log('res-links', missing.length + ' éléments dans la liste des figures n\'ont pas de lien.');
            }
          } catch (err) {
            log('res-links', 'Erreur: ' + err.message);
            console.error(err);
          }
        };
      }

      // bouton - init (optionnel, si présent dans ton HTML)
      const initBtn = document.getElementById('btn-init');
      if (initBtn) {
        initBtn.onclick = async function() {
          try {
            await Word.run(async (context) => {
              const body = context.document.body;
              body.load('text');
              await context.sync();
              if (body.text && body.text.trim().length > 0) {
                const keep = confirm("Le document contient déjà du texte. Effacer pour créer le document d'exemple ?");
                if (!keep) return;
                body.clear();
              }

              body.insertParagraph('Exercice : Insérer plusieurs images à différents emplacements, ajouter une légende sous chaque image (centrée, numérotée automatiquement), et générer une liste des figures.', Word.InsertLocation.end);
              body.insertParagraph('\n--- Page 1 ---\n', Word.InsertLocation.end);
              body.insertParagraph('Instructions répétées pour remplir le document...', Word.InsertLocation.end);
              body.insertParagraph('\n--- Page 2 ---\n', Word.InsertLocation.end);
              body.insertParagraph('Placez des images manuellement à différents emplacements, puis ajoutez une légende sous chaque image de la forme "Figure 1 : ..."', Word.InsertLocation.end);

              await context.sync();
              log('console', 'Document d\'exemple créé (partiel). Ajoutez des images manuelles puis utilisez les vérifications.');
            });
          } catch (err) {
            log('console', 'Erreur initialisation: ' + err.message);
            console.error(err);
          }
        };
      }

      log('console', 'Complément prêt à l\'emploi !');
    }); // end DOMContentLoaded
  } // end if host
}); // end Office.onReady
