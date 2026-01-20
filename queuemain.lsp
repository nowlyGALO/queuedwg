;;; ============================================================
;;; PROCESSADOR DE FILA (controle.xlsx + config.ini)
;;; - SEM caminhos fixos (portátil para outros PCs)
;;; - Lê config.ini gerado pelo Excel (MODE/MATCH/OLD/NEW/MODEL/PAPER/TAG/BLOCK)
;;; - Lê controle.xlsx (Coluna A = Nome, Coluna B = Caminho)
;;; - Abre cada DWG, altera (TEXT/MTEXT ou ATTRIB), REGEN, SAVE
;;; - NÃO FECHA os arquivos (fechamento manual)
;;; - Gera log detalhado na mesma pasta do config.ini
;;; ============================================================

(vl-load-com)

;;; ----------------------------
;;; Util: data/hora para log
;;; ----------------------------
(defun _now ()
  (menucmd "M=$(edtime,$(getvar,date),yyyy-mm-dd hh:MM:ss)")
)

(defun _log-open (logPath / arq)
  (setq arq (open logPath "a"))
  arq
)

(defun _log-line (arq msg)
  (if arq
    (write-line (strcat (_now) " " msg) arq)
  )
)

(defun _log-close (arq)
  (if arq (close arq))
)

;;; ----------------------------
;;; Util: strings / variantes
;;; ----------------------------
(defun _safe-to-str (v)
  (cond
    ((null v) "")
    ((= (type v) 'VARIANT) (_safe-to-str (vlax-variant-value v)))
    ((= (type v) 'STR) v)
    (t (vl-princ-to-string v))
  )
)

(defun _trim (s)
  (if s (vl-string-trim " \t\r\n" s) "")
)

(defun _strcase (s)
  (if s (strcase s) "")
)

(defun _contains (hay needle)
  (if (and hay needle)
    (not (null (vl-string-search needle hay)))
    nil
  )
)

;;; ----------------------------
;;; Util: caminho / pasta
;;; ----------------------------
(defun _path-dir (full / p s)
  (setq s (vl-string-translate "/" "\\" full))
  (setq p (strlen s))
  (while (and (> p 0) (/= (substr s p 1) "\\"))
    (setq p (1- p))
  )
  (if (> p 0)
    (substr s 1 p)
    ""
  )
)

(defun _path-join (dir file)
  (cond
    ((= dir "") file)
    ((= (substr dir (strlen dir) 1) "\\") (strcat dir file))
    (t (strcat dir "\\" file))
  )
)

;;; ----------------------------
;;; Ler config.ini (CHAVE=VALOR)
;;; Retorna lista assoc: (("MODE" . "TEXT") ...)
;;; ----------------------------
(defun _read-config-ini (path / arq linha pos chave valor cfg)
  (setq cfg '())
  (if (setq arq (open path "r"))
    (progn
      (while (setq linha (read-line arq))
        (setq linha (_trim linha))
        (if (and (/= linha "") (setq pos (vl-string-search "=" linha)))
          (progn
            (setq chave (_strcase (substr linha 1 pos)))
            (setq valor (substr linha (+ pos 2)))
            (setq cfg (cons (cons chave value) cfg)) ;; placeholder fixed below
          )
        )
      )
      (close arq)
    )
  )
  cfg
)

;;; *** CORREÇÃO: LISP não permite variável "value" não definida. Reescrevendo função corretamente:
(defun _read-config-ini (path / arq linha pos chave valor cfg)
  (setq cfg '())
  (if (setq arq (open path "r"))
    (progn
      (while (setq linha (read-line arq))
        (setq linha (_trim linha))
        (if (and (/= linha "") (setq pos (vl-string-search "=" linha)))
          (progn
            (setq chave (_strcase (substr linha 1 pos)))
            (setq valor (substr linha (+ pos 2)))
            (setq cfg (cons (cons chave valor) cfg))
          )
        )
      )
      (close arq)
    )
  )
  cfg
)

(defun _cfg-get (cfg key / a)
  (setq a (assoc (_strcase key) cfg))
  (if a (cdr a) "")
)

(defun _cfg-bool (cfg key)
  (= (_cfg-get cfg key) "1")
)

;;; ----------------------------
;;; Excel: ler controle.xlsx (A nome, B caminho)
;;; Retorna lista de caminhos
;;; ----------------------------
(defun _xl-get (ws addr / rng val)
  (setq rng (vlax-get-property ws 'Range addr))
  (setq val (vlax-get-property rng 'Value))
  (_safe-to-str val)
)

(defun _read-controle-xlsx (xlsxPath logArq / xl wbs wb ws linha nome caminho lst usedRows r)
  (setq lst '())

  (_log-line logArq (strcat "Abrindo Excel (somente leitura): " xlsxPath))

  (setq xl (vlax-create-object "Excel.Application"))
  (vlax-put-property xl 'Visible :vlax-false)
  (setq wbs (vlax-get-property xl 'Workbooks))

  ;; Open(Filename, UpdateLinks, ReadOnly)
  (setq wb (vl-catch-all-apply 'vlax-invoke-method (list wbs 'Open xlsxPath nil :vlax-true)))
  (if (vl-catch-all-error-p wb)
    (progn
      (_log-line logArq "ERRO: Falha ao abrir o controle.xlsx no Excel.")
      (vl-catch-all-apply 'vlax-invoke-method (list xl 'Quit))
      (vlax-release-object wbs)
      (vlax-release-object xl)
      '()
    )
    (progn
      (setq wb (vlax-variant-value wb))
      (setq ws (vlax-get-property wb 'ActiveSheet))

      ;; tenta pegar UsedRange.Rows.Count (para log)
      (setq usedRows 0)
      (setq r (vl-catch-all-apply 'vlax-get-property (list (vlax-get-property (vlax-get-property ws 'UsedRange) 'Rows) 'Count)))
      (if (not (vl-catch-all-error-p r))
        (setq usedRows (vlax-variant-value r))
      )
      (_log-line logArq (strcat "Linhas usadas na planilha (aprox): " (itoa usedRows)))

      (setq linha 2)
      (while T
        (setq nome (_trim (_xl-get ws (strcat "A" (itoa linha)))))
        (setq caminho (_trim (_xl-get ws (strcat "B" (itoa linha)))))

        (if (or (= nome "") (= caminho ""))
          (progn
            (_log-line logArq (strcat "Fim da lista na linha " (itoa linha) " (celula vazia)."))
            (setq linha nil)
            (quit)
          )
        )

        (setq lst (cons caminho lst))
        (setq linha (1+ linha))
      )

      ;; Fecha Excel
      (_log-line logArq "Fechando Excel (somente leitura).")
      (vl-catch-all-apply 'vlax-invoke-method (list wb 'Close :vlax-false))
      (vl-catch-all-apply 'vlax-invoke-method (list xl 'Quit))

      (mapcar 'vlax-release-object (list ws wb wbs xl))

      (reverse lst)
    )
  )
)

;;; ----------------------------
;;; TEXT/MTEXT: get/set safe
;;; ----------------------------
(defun _set-textstring-safe (obj new / r)
  (setq r (vl-catch-all-apply 'vla-put-TextString (list obj new)))
  (not (vl-catch-all-error-p r))
)

(defun _get-textstring-safe (obj / r)
  (setq r (vl-catch-all-apply 'vla-get-TextString (list obj)))
  (if (vl-catch-all-error-p r) nil r)
)

;;; ----------------------------
;;; REGEN safe
;;; ----------------------------
(defun _try-regen (doc logArq / r)
  ;; acAllViewports = 2
  (setq r (vl-catch-all-apply 'vla-Regen (list doc 2)))
  (if (vl-catch-all-error-p r)
    (progn (_log-line logArq "ERRO: REGEN falhou") nil)
    (progn (_log-line logArq "REGEN OK") T)
  )
)

;;; ----------------------------
;;; Processar TEXT/MTEXT em uma "space"
;;; matchMode: "EQUAL" ou "CONTAINS"
;;; ----------------------------
(defun _process-entity-text (obj matchMode old new / ts changed exact contains newts)
  (setq changed 0)
  (setq ts (_get-textstring-safe obj))
  (if ts
    (progn
      (setq exact (= ts old))
      (setq contains (_contains ts old))

      (cond
        ((= matchMode "EQUAL")
          (if exact
            (if (_set-textstring-safe obj new) (setq changed 1))
          )
        )
        ((= matchMode "CONTAINS")
          (if contains
            (progn
              (setq newts (vl-string-subst new old ts))
              (if (_set-textstring-safe obj newts) (setq changed 1))
            )
          )
        )
        (T
          ;; fallback: EQUAL
          (if exact
            (if (_set-textstring-safe obj new) (setq changed 1))
          )
        )
      )
    )
  )
  changed
)

(defun _process-space-text (space matchMode old new / obj oname chg total)
  (setq total 0)
  (vlax-for obj space
    (setq oname (vla-get-ObjectName obj))
    (if (or (= oname "AcDbText") (= oname "AcDbMText"))
      (progn
        (setq chg (_process-entity-text obj matchMode old new))
        (if (> chg 0) (setq total (+ total chg)))
      )
    )
  )
  total
)

;;; ----------------------------
;;; ATTRIB: processar atributos em BlockReference
;;; filtros opcionais:
;;; - tagFilter: "" ou TAG
;;; - blockFilter: "" ou nome do bloco
;;; ----------------------------
(defun _blk-effective-name (blk / r)
  ;; em alguns casos existe EffectiveName; se falhar, usa Name
  (setq r (vl-catch-all-apply 'vla-get-EffectiveName (list blk)))
  (if (vl-catch-all-error-p r)
    (_safe-to-str (vl-catch-all-apply 'vla-get-Name (list blk)))
    (_safe-to-str r)
  )
)

(defun _process-attr-one (att matchMode old new / tag val changed exact contains newval)
  (setq changed 0)

  (setq tag (_safe-to-str (vl-catch-all-apply 'vla-get-TagString (list att))))
  (setq val (_safe-to-str (vl-catch-all-apply 'vla-get-TextString (list att))))

  (cond
    ((= matchMode "EQUAL")
      (if (= val old)
        (progn
          (if (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-put-TextString (list att new))))
            (setq changed 1)
          )
        )
      )
    )
    ((= matchMode "CONTAINS")
      (if (_contains val old)
        (progn
          (setq newval (vl-string-subst new old val))
          (if (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-put-TextString (list att newval))))
            (setq changed 1)
          )
        )
      )
    )
    (T
      (if (= val old)
        (progn
          (if (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-put-TextString (list att new))))
            (setq changed 1)
          )
        )
      )
    )
  )
  changed
)

(defun _process-block-attrs (blk matchMode old new tagFilter blockFilter / changed hasAtt arrAtt v i att tag bn)
  (setq changed 0)

  ;; filtro por nome do bloco (opcional)
  (setq bn (_strcase (_trim (_blk-effective-name blk))))
  (if (and (/= (_trim blockFilter) "") (/= bn (_strcase (_trim blockFilter))))
    (progn
      0
    )
    (progn
      (setq hasAtt (vl-catch-all-apply 'vla-get-HasAttributes (list blk)))
      (if (or (vl-catch-all-error-p hasAtt) (= (vlax-variant-value hasAtt) :vlax-false))
        0
        (progn
          (setq arrAtt (vl-catch-all-apply 'vla-GetAttributes (list blk)))
          (if (vl-catch-all-error-p arrAtt)
            0
            (progn
              (setq v (vlax-variant-value arrAtt))
              (setq i 0)
              (while (< i (vlax-safearray-get-u-bound v 1))
                (setq att (vlax-safearray-get-element v i))

                ;; filtro por TAG (opcional)
                (setq tag (_strcase (_trim (_safe-to-str (vl-catch-all-apply 'vla-get-TagString (list att))))))
                (if (or (= (_trim tagFilter) "") (= tag (_strcase (_trim tagFilter))))
                  (setq changed (+ changed (_process-attr-one att matchMode old new)))
                )

                (setq i (1+ i))
              )
              changed
            )
          )
        )
      )
    )
  )
)

(defun _process-space-attrib (space matchMode old new tagFilter blockFilter / obj oname total)
  (setq total 0)
  (vlax-for obj space
    (setq oname (vla-get-ObjectName obj))
    (if (= oname "AcDbBlockReference")
      (setq total (+ total (_process-block-attrs obj matchMode old new tagFilter blockFilter)))
    )
  )
  total
)

;;; ----------------------------
;;; Abrir/ativar doc, processar, salvar (não fecha)
;;; ----------------------------
(defun _open-doc (path / docs doc r)
  (setq docs (vla-get-Documents (vlax-get-acad-object)))
  (setq r (vl-catch-all-apply 'vla-open (list docs path)))
  (if (vl-catch-all-error-p r)
    nil
    (progn
      (setq doc r)
      (vl-catch-all-apply 'vla-activate (list doc))
      doc
    )
  )
)

(defun _save-doc (doc logArq / r)
  (setq r (vl-catch-all-apply 'vla-save (list doc)))
  (if (vl-catch-all-error-p r)
    (progn (_log-line logArq "ERRO: falha ao salvar (vla-save).") nil)
    (progn (_log-line logArq "SAVE OK") T)
  )
)

;;; ============================================================
;;; COMANDO PRINCIPAL
;;; ============================================================
(defun c:PROCESSA_FILA_CONFIG ( / cfgPath baseDir xlsxPath logPath logArq cfg mode match old new applyModel applyPaper tagFilter blockFilter dwgList totalFiles totalChangedFiles totalChanges i path doc chgM chgP chgTotal)

  ;; 1) Escolher config.ini (portátil)
  (setq cfgPath (getfiled "Selecione o config.ini gerado pelo Excel" "" "ini" 16))
  (if (or (null cfgPath) (= cfgPath ""))
    (progn (prompt "\nCancelado.") (princ) (exit))
  )

  (setq baseDir (_path-dir cfgPath))
  (setq xlsxPath (_path-join baseDir "controle.xlsx"))
  (setq logPath (_path-join baseDir "log_processamento.txt"))

  ;; 2) Abrir log (zera no início)
  (if (findfile logPath) (vl-file-delete logPath))
  (setq logArq (_log-open logPath))

  (_log-line logArq "================ INICIO PROCESSAMENTO ================")
  (_log-line logArq (strcat "Config: " cfgPath))
  (_log-line logArq (strcat "Controle: " xlsxPath))
  (_log-line logArq (strcat "Pasta Base: " baseDir))

  ;; 3) Ler config.ini
  (setq cfg (_read-config-ini cfgPath))
  (setq mode (_strcase (_trim (_cfg-get cfg "MODE"))))      ;; TEXT / ATTRIB
  (setq match (_strcase (_trim (_cfg-get cfg "MATCH"))))    ;; EQUAL / CONTAINS
  (setq old (_cfg-get cfg "OLD"))
  (setq new (_cfg-get cfg "NEW"))
  (setq applyModel (_cfg-bool cfg "MODEL"))
  (setq applyPaper (_cfg-bool cfg "PAPER"))
  (setq tagFilter (_cfg-get cfg "TAG"))
  (setq blockFilter (_cfg-get cfg "BLOCK"))

  (_log-line logArq (strcat "MODE=" mode " | MATCH=" match))
  (_log-line logArq (strcat "OLD=" old " | NEW=" new))
  (_log-line logArq (strcat "MODEL=" (if applyModel "1" "0") " | PAPER=" (if applyPaper "1" "0")))
  (_log-line logArq (strcat "TAG=" tagFilter " | BLOCK=" blockFilter))

  (if (not (findfile xlsxPath))
    (progn
      (_log-line logArq "ERRO: controle.xlsx nao encontrado na mesma pasta do config.ini.")
      (_log-close logArq)
      (alert "ERRO: controle.xlsx não encontrado na mesma pasta do config.ini.")
      (princ)
      (exit)
    )
  )

  ;; 4) Ler controle.xlsx (lista de caminhos)
  (setq dwgList (_read-controle-xlsx xlsxPath logArq))
  (if (or (null dwgList) (= (length dwgList) 0))
    (progn
      (_log-line logArq "ERRO: lista de DWGs vazia.")
      (_log-close logArq)
      (alert "ERRO: lista de DWGs vazia no controle.xlsx.")
      (princ)
      (exit)
    )
  )

  ;; 5) Processar cada DWG
  (setq totalFiles 0)
  (setq totalChangedFiles 0)
  (setq totalChanges 0)

  (foreach path dwgList
    (setq totalFiles (1+ totalFiles))
    (_log-line logArq "------------------------------------------------------")
    (_log-line logArq (strcat "Arquivo " (itoa totalFiles) ": " path))

    (if (not (findfile path))
      (_log-line logArq "ERRO: arquivo nao encontrado (findfile).")
      (progn
        (_log-line logArq "Abrindo DWG...")
        (setq doc (_open-doc path))

        (if (null doc)
          (_log-line logArq "ERRO: falha ao abrir DWG (vla-open).")
          (progn
            (setq chgM 0)
            (setq chgP 0)

            (cond
              ((= mode "TEXT")
                (_log-line logArq "Modo TEXT/MTEXT.")
                (if applyModel
                  (progn
                    (_log-line logArq "Varredura ModelSpace (TEXT/MTEXT)...")
                    (setq chgM (_process-space-text (vla-get-ModelSpace doc) match old new))
                    (_log-line logArq (strcat "Alteracoes ModelSpace: " (itoa chgM)))
                  )
                  (_log-line logArq "ModelSpace desativado no config.")
                )
                (if applyPaper
                  (progn
                    (_log-line logArq "Varredura PaperSpace (TEXT/MTEXT)...")
                    (setq chgP (_process-space-text (vla-get-PaperSpace doc) match old new))
                    (_log-line logArq (strcat "Alteracoes PaperSpace: " (itoa chgP)))
                  )
                  (_log-line logArq "PaperSpace desativado no config.")
                )
              )

              ((= mode "ATTRIB")
                (_log-line logArq "Modo ATTRIB (atributos em blocos).")
                (if applyModel
                  (progn
                    (_log-line logArq "Varredura ModelSpace (ATTRIB)...")
                    (setq chgM (_process-space-attrib (vla-get-ModelSpace doc) match old new tagFilter blockFilter))
                    (_log-line logArq (strcat "Alteracoes ModelSpace: " (itoa chgM)))
                  )
                  (_log-line logArq "ModelSpace desativado no config.")
                )
                (if applyPaper
                  (progn
                    (_log-line logArq "Varredura PaperSpace (ATTRIB)...")
                    (setq chgP (_process-space-attrib (vla-get-PaperSpace doc) match old new tagFilter blockFilter))
                    (_log-line logArq (strcat "Alteracoes PaperSpace: " (itoa chgP)))
                  )
                  (_log-line logArq "PaperSpace desativado no config.")
                )
              )

              (T
                (_log-line logArq "ERRO: MODE invalido no config.ini (use TEXT ou ATTRIB).")
              )
            )

            (setq chgTotal (+ chgM chgP))
            (_log-line logArq (strcat "RESUMO: total alterado = " (itoa chgTotal)))

            (if (> chgTotal 0)
              (progn
                (setq totalChangedFiles (1+ totalChangedFiles))
                (setq totalChanges (+ totalChanges chgTotal))
                (_log-line logArq "Executando REGEN apos alteracao...")
                (_try-regen doc logArq)
              )
              (_log-line logArq "Sem alteracao -> sem REGEN.")
            )

            (_log-line logArq "Salvando DWG (sem fechar)...")
            (_save-doc doc logArq)

            (_log-line logArq "DWG mantido aberto (fechamento manual).")
          )
        )
      )
    )
  )

  ;; 6) Final do log
  (_log-line logArq "================ RESUMO GERAL =================")
  (_log-line logArq (strcat "Total arquivos na fila: " (itoa totalFiles)))
  (_log-line logArq (strcat "Arquivos com alteracao: " (itoa totalChangedFiles)))
  (_log-line logArq (strcat "Total de alteracoes: " (itoa totalChanges)))
  (_log-line logArq "================ FIM PROCESSAMENTO ================")

  (_log-close logArq)

  (alert (strcat
    "Processamento concluido (SEM FECHAR arquivos)." "\n\n"
    "Total arquivos: " (itoa totalFiles) "\n"
    "Arquivos alterados: " (itoa totalChangedFiles) "\n"
    "Total alteracoes: " (itoa totalChanges) "\n\n"
    "Log: " logPath
  ))

  (princ)
)
