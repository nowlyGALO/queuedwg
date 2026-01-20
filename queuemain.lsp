;;; ============================================================
;;; PROCESSA FILA (controle.xlsx + config.ini) - ROBUSTO
;;; - Portátil: você escolhe o config.ini
;;; - controle.xlsx deve estar na MESMA PASTA do config.ini
;;; - Lê config.ini: MODE/MATCH/OLD/NEW/MODEL/PAPER/TAG/BLOCK/CLOSE
;;; - Lê controle.xlsx: Col A=Nome, Col B=Caminho
;;; - Abre cada DWG, altera (TEXT/MTEXT ou ATTRIB), REGEN, SAVE
;;; - Fecha os arquivos se CLOSE=1 (robusto, com fallback)
;;; - Log na mesma pasta do config.ini
;;; ============================================================

(vl-load-com)

;;; ----------------------------
;;; Log
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

(defun _err-msg (e)
  (if (vl-catch-all-error-p e)
    (vl-catch-all-error-message e)
    ""
  )
)

;;; ----------------------------
;;; Strings / Variants
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

(defun _itoa-safe (n)
  (if (numberp n) (itoa n) "0")
)

;;; ----------------------------
;;; Paths
;;; ----------------------------
(defun _path-dir (full / s p)
  (setq s (vl-string-translate "/" "\\" full))
  (setq p (strlen s))
  (while (and (> p 0) (/= (substr s p 1) "\\"))
    (setq p (1- p))
  )
  (if (> p 0) (substr s 1 p) "")
)

(defun _path-join (dir file)
  (cond
    ((= dir "") file)
    ((= (substr dir (strlen dir) 1) "\\") (strcat dir file))
    (t (strcat dir "\\" file))
  )
)

;;; ----------------------------
;;; Config.ini
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
            (setq valor (_trim (substr linha (+ pos 2))))
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

(defun _cfg-bool (cfg key / val)
  (setq val (_strcase (_trim (_cfg-get cfg key))))
  (or (= val "1") (= val "TRUE") (= val "YES") (= val "SIM"))
)

;;; ----------------------------
;;; Doc helpers
;;; ----------------------------
(defun _doc-open-p (doc / docs item isOpen)
  (setq isOpen nil)
  (setq docs (vla-get-Documents (vlax-get-acad-object)))
  (vlax-for item docs
    (if (= item doc)
      (setq isOpen T)
    )
  )
  isOpen
)

;;; ----------------------------
;;; Excel read
;;; ----------------------------
(defun _xl-get (ws addr / rng val)
  (setq rng (vlax-get-property ws 'Range addr))
  (setq val (vlax-get-property rng 'Value))
  ;; libera Range (evita vazamento COM)
  (if rng (vlax-release-object rng))
  (_safe-to-str val)
)

(defun _read-controle-xlsx (xlsxPath logArq / xl wbs wb ws linha nome caminho lst done)
  (setq lst '())
  (setq done nil)

  (_log-line logArq (strcat "Abrindo Excel (somente leitura): " xlsxPath))

  (setq xl (vlax-create-object "Excel.Application"))
  (vlax-put-property xl 'Visible :vlax-false)
  (setq wbs (vlax-get-property xl 'Workbooks))

  ;; Open(Filename, UpdateLinks, ReadOnly)
  (setq wb (vl-catch-all-apply 'vlax-invoke-method (list wbs 'Open xlsxPath nil :vlax-true)))
  (if (vl-catch-all-error-p wb)
    (progn
      (_log-line logArq (strcat "ERRO: Falha ao abrir o controle.xlsx no Excel. " (_err-msg wb)))
      (vl-catch-all-apply 'vlax-invoke-method (list xl 'Quit))
      (mapcar 'vlax-release-object (list wbs xl))
      '()
    )
    (progn
      ;; Mantém seu comportamento original: usar ActiveSheet
      (setq ws (vlax-get-property wb 'ActiveSheet))

      (setq linha 2)

      (while (not done)
        (setq nome (_trim (_xl-get ws (strcat "A" (_itoa-safe linha)))))
        (setq caminho (_trim (_xl-get ws (strcat "B" (_itoa-safe linha)))))

        (if (or (= nome "") (= caminho ""))
          (progn
            (_log-line logArq (strcat "Fim da lista na linha " (_itoa-safe linha) " (A ou B vazia)."))
            (setq done T)
          )
          (progn
            (setq lst (cons caminho lst))
            (setq linha (1+ linha))
          )
        )
      )

      (_log-line logArq "Fechando Excel (somente leitura).")
      (vl-catch-all-apply 'vlax-invoke-method (list wb 'Close :vlax-false))
      (vl-catch-all-apply 'vlax-invoke-method (list xl 'Quit))
      (mapcar 'vlax-release-object (list ws wb wbs xl))

      (reverse lst)
    )
  )
)

;;; ----------------------------
;;; TEXT/MTEXT get/set
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
;;; REGEN / SAVE / CLOSE
;;; ----------------------------
(defun _try-regen (doc logArq / r)
  ;; acAllViewports = 2
  (setq r (vl-catch-all-apply 'vla-Regen (list doc 2)))
  (if (vl-catch-all-error-p r)
    (progn (_log-line logArq (strcat "ERRO: REGEN falhou. " (_err-msg r))) nil)
    (progn (_log-line logArq "REGEN OK") T)
  )
)

(defun _save-doc (doc logArq / r)
  (setq r (vl-catch-all-apply 'vla-save (list doc)))
  (if (vl-catch-all-error-p r)
    (progn (_log-line logArq (strcat "ERRO: falha ao salvar (vla-save). " (_err-msg r))) nil)
    (progn (_log-line logArq "SAVE OK") T)
  )
)

(defun _close-doc (doc logArq / r oldFileDia oldCmdDia)
  ;; 1) tenta vla-close sem parâmetro
  (_log-line logArq "Tentando fechar via vla-close (sem parametro)...")
  (setq r (vl-catch-all-apply 'vla-close (list doc)))
  (if (not (vl-catch-all-error-p r))
    (progn (_log-line logArq "CLOSE OK (vla-close sem parametro)") T)
    (progn
      (_log-line logArq (strcat "Falhou vla-close sem parametro: " (_err-msg r)))

      ;; 2) tenta vla-close com SaveChanges = True (algumas versões preferem/precisam)
      (_log-line logArq "Tentando fechar via vla-close (:vlax-true)...")
      (setq r (vl-catch-all-apply 'vla-close (list doc :vlax-true)))
      (if (not (vl-catch-all-error-p r))
        (progn (_log-line logArq "CLOSE OK (vla-close :vlax-true)") T)
        (progn
          (_log-line logArq (strcat "Falhou vla-close :vlax-true: " (_err-msg r)))

          ;; 3) fallback por comando (pode contornar interferência de add-ons)
          (_log-line logArq "Fallback: fechando via comando _.CLOSE (sem dialogos)...")
          (setq oldFileDia (getvar "FILEDIA"))
          (setq oldCmdDia  (getvar "CMDDIA"))
          (setvar "FILEDIA" 0)
          (setvar "CMDDIA"  0)

          ;; garante documento ativo
          (vl-catch-all-apply 'vla-activate (list doc))

          ;; fecha confirmando salvamento (já salvamos, mas confirma)
          ;; comando: CLOSE -> pede salvar? (Y/N). Usamos "Y".
          (vl-catch-all-apply 'vl-cmdf (list "_.CLOSE" "_Y"))

          (setvar "FILEDIA" oldFileDia)
          (setvar "CMDDIA"  oldCmdDia)

          (if (_doc-open-p doc)
            (progn
              (_log-line logArq "ATENCAO: fallback _.CLOSE executado, mas o desenho ainda esta aberto.")
              nil
            )
            (progn
              (_log-line logArq "Fallback _.CLOSE executado e desenho fechado.")
              T
            )
          )
        )
      )
    )
  )
)

;;; ----------------------------
;;; TEXT/MTEXT processing
;;; matchMode: EQUAL | CONTAINS
;;; ----------------------------
(defun _process-entity-text (obj matchMode old new / ts changed newts)
  (setq changed 0)
  (setq ts (_get-textstring-safe obj))
  (if ts
    (cond
      ((= matchMode "EQUAL")
        (if (= ts old)
          (if (_set-textstring-safe obj new) (setq changed 1))
        )
      )
      ((= matchMode "CONTAINS")
        (if (_contains ts old)
          (progn
            (setq newts (vl-string-subst new old ts))
            (if (_set-textstring-safe obj newts) (setq changed 1))
          )
        )
      )
      (T
        (if (= ts old)
          (if (_set-textstring-safe obj new) (setq changed 1))
        )
      )
    )
  )
  changed
)

(defun _process-space-text (space matchMode old new / obj oname total)
  (setq total 0)
  (vlax-for obj space
    (setq oname (vla-get-ObjectName obj))
    (if (or (= oname "AcDbText") (= oname "AcDbMText"))
      (setq total (+ total (_process-entity-text obj matchMode old new)))
    )
  )
  total
)

;;; ----------------------------
;;; ATTRIB processing
;;; ----------------------------
(defun _blk-effective-name (blk / r)
  (setq r (vl-catch-all-apply 'vla-get-EffectiveName (list blk)))
  (if (vl-catch-all-error-p r)
    (_safe-to-str (vl-catch-all-apply 'vla-get-Name (list blk)))
    (_safe-to-str r)
  )
)

(defun _process-attr-one (att matchMode old new / val changed newval)
  (setq changed 0)
  (setq val (_safe-to-str (vl-catch-all-apply 'vla-get-TextString (list att))))

  (cond
    ((= matchMode "EQUAL")
      (if (= val old)
        (if (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-put-TextString (list att new))))
          (setq changed 1)
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
        (if (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-put-TextString (list att new))))
          (setq changed 1)
        )
      )
    )
  )
  changed
)

(defun _process-block-attrs (blk matchMode old new tagFilter blockFilter / changed bn hasAtt arrAtt sa lb ub i att tag)
  (setq changed 0)

  ;; filtro por bloco (opcional)
  (setq bn (_strcase (_trim (_blk-effective-name blk))))
  (if (and (/= (_trim blockFilter) "") (/= bn (_strcase (_trim blockFilter))))
    0
    (progn
      (setq hasAtt (vl-catch-all-apply 'vla-get-HasAttributes (list blk)))
      (if (or (vl-catch-all-error-p hasAtt) (= (vlax-variant-value hasAtt) :vlax-false))
        0
        (progn
          (setq arrAtt (vl-catch-all-apply 'vla-GetAttributes (list blk)))
          (if (vl-catch-all-error-p arrAtt)
            0
            (progn
              (setq sa (vlax-variant-value arrAtt))

              ;; bounds protegidos
              (setq lb (vl-catch-all-apply 'vlax-safearray-get-l-bound (list sa 1)))
              (setq ub (vl-catch-all-apply 'vlax-safearray-get-u-bound (list sa 1)))

              (if (or (vl-catch-all-error-p lb) (vl-catch-all-error-p ub))
                0
                (progn
                  (setq lb (vlax-variant-value lb))
                  (setq ub (vlax-variant-value ub))

                  (if (or (null lb) (null ub))
                    0
                    (progn
                      (setq i lb)
                      (while (<= i ub)
                        (setq att (vlax-safearray-get-element sa i))

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
;;; Abrir/ativar doc
;;; ----------------------------
(defun _open-doc (path / docs r)
  (setq docs (vla-get-Documents (vlax-get-acad-object)))
  (setq r (vl-catch-all-apply 'vla-open (list docs path)))
  (if (vl-catch-all-error-p r)
    nil
    (progn
      (vl-catch-all-apply 'vla-activate (list r))
      r
    )
  )
)

;;; ============================================================
;;; COMANDO PRINCIPAL
;;; ============================================================
(defun c:PROCESSA_FILA_CONFIG
  ( / cfgPath baseDir xlsxPath logPath logArq cfg
    mode match old new applyModel applyPaper tagFilter blockFilter closeAfter
    dwgList totalFiles totalChangedFiles totalChanges
    path doc chgM chgP chgTotal)

  (setq cfgPath (getfiled "Selecione o config.ini gerado pelo Excel" "" "ini" 16))
  (if (or (null cfgPath) (= cfgPath ""))
    (progn (prompt "\nCancelado.") (princ))
    (progn
      (setq baseDir (_path-dir cfgPath))
      (setq xlsxPath (_path-join baseDir "controle.xlsx"))
      (setq logPath (_path-join baseDir "log_processamento.txt"))

      (if (findfile logPath) (vl-file-delete logPath))
      (setq logArq (_log-open logPath))

      (_log-line logArq "================ INICIO PROCESSAMENTO ================")
      (_log-line logArq (strcat "Config: " cfgPath))
      (_log-line logArq (strcat "Controle: " xlsxPath))
      (_log-line logArq (strcat "Pasta Base: " baseDir))

      ;; Ler config
      (setq cfg (_read-config-ini cfgPath))
      (setq mode (_strcase (_trim (_cfg-get cfg "MODE"))))
      (setq match (_strcase (_trim (_cfg-get cfg "MATCH"))))
      (setq old (_cfg-get cfg "OLD"))
      (setq new (_cfg-get cfg "NEW"))
      (setq applyModel (_cfg-bool cfg "MODEL"))
      (setq applyPaper (_cfg-bool cfg "PAPER"))
      (setq tagFilter (_cfg-get cfg "TAG"))
      (setq blockFilter (_cfg-get cfg "BLOCK"))
      (setq closeAfter (_cfg-bool cfg "CLOSE"))

      (_log-line logArq (strcat "MODE=" mode " | MATCH=" match))
      (_log-line logArq (strcat "OLD=" old " | NEW=" new))
      (_log-line logArq (strcat "MODEL=" (if applyModel "1" "0") " | PAPER=" (if applyPaper "1" "0")))
      (_log-line logArq (strcat "TAG=" tagFilter " | BLOCK=" blockFilter))
      (_log-line logArq (strcat "CLOSE=" (if closeAfter "1" "0")))

      (if (not (findfile xlsxPath))
        (progn
          (_log-line logArq "ERRO: controle.xlsx nao encontrado na mesma pasta do config.ini.")
          (_log-line logArq "================ FIM (ERRO) =================")
          (_log-close logArq)
          (alert "ERRO: controle.xlsx não encontrado na mesma pasta do config.ini.")
        )
        (progn
          ;; Ler lista do Excel
          (setq dwgList (_read-controle-xlsx xlsxPath logArq))

          (if (or (null dwgList) (= (length dwgList) 0))
            (progn
              (_log-line logArq "ERRO: lista de DWGs vazia.")
              (_log-line logArq "================ FIM (ERRO) =================")
              (_log-close logArq)
              (alert "ERRO: lista de DWGs vazia no controle.xlsx.")
            )
            (progn
              ;; Processar
              (setq totalFiles 0)
              (setq totalChangedFiles 0)
              (setq totalChanges 0)

              (foreach path dwgList
                (setq totalFiles (1+ totalFiles))
                (_log-line logArq "------------------------------------------------------")
                (_log-line logArq (strcat "Arquivo " (_itoa-safe totalFiles) ": " path))

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
                                (_log-line logArq (strcat "Alteracoes ModelSpace: " (_itoa-safe chgM)))
                              )
                              (_log-line logArq "ModelSpace desativado no config.")
                            )
                            (if applyPaper
                              (progn
                                (_log-line logArq "Varredura PaperSpace (TEXT/MTEXT)...")
                                (setq chgP (_process-space-text (vla-get-PaperSpace doc) match old new))
                                (_log-line logArq (strcat "Alteracoes PaperSpace: " (_itoa-safe chgP)))
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
                                (_log-line logArq (strcat "Alteracoes ModelSpace: " (_itoa-safe chgM)))
                              )
                              (_log-line logArq "ModelSpace desativado no config.")
                            )
                            (if applyPaper
                              (progn
                                (_log-line logArq "Varredura PaperSpace (ATTRIB)...")
                                (setq chgP (_process-space-attrib (vla-get-PaperSpace doc) match old new tagFilter blockFilter))
                                (_log-line logArq (strcat "Alteracoes PaperSpace: " (_itoa-safe chgP)))
                              )
                              (_log-line logArq "PaperSpace desativado no config.")
                            )
                          )

                          (T
                            (_log-line logArq "ERRO: MODE invalido no config.ini (use TEXT ou ATTRIB).")
                          )
                        )

                        (setq chgTotal (+ chgM chgP))
                        (_log-line logArq (strcat "RESUMO: total alterado = " (_itoa-safe chgTotal)))

                        (if (> chgTotal 0)
                          (progn
                            (setq totalChangedFiles (1+ totalChangedFiles))
                            (setq totalChanges (+ totalChanges chgTotal))
                            (_log-line logArq "Executando REGEN apos alteracao...")
                            (_try-regen doc logArq)
                          )
                          (_log-line logArq "Sem alteracao -> sem REGEN.")
                        )

                        (_log-line logArq "Salvando DWG...")
                        (_save-doc doc logArq)

                        (if closeAfter
                          (progn
                            (_log-line logArq "Fechando DWG (CLOSE=1)...")
                            (_close-doc doc logArq)
                          )
                          (_log-line logArq "DWG mantido aberto (CLOSE=0).")
                        )
                      )
                    )
                  )
                )
              )

              (_log-line logArq "================ RESUMO GERAL =================")
              (_log-line logArq (strcat "Total arquivos na fila: " (_itoa-safe totalFiles)))
              (_log-line logArq (strcat "Arquivos com alteracao: " (_itoa-safe totalChangedFiles)))
              (_log-line logArq (strcat "Total de alteracoes: " (_itoa-safe totalChanges)))
              (_log-line logArq "================ FIM PROCESSAMENTO ================")

              (_log-close logArq)

              (alert (strcat
                "Processamento concluido." "\n\n"
                "Total arquivos: " (_itoa-safe totalFiles) "\n"
                "Arquivos alterados: " (_itoa-safe totalChangedFiles) "\n"
                "Total alteracoes: " (_itoa-safe totalChanges) "\n\n"
                "Log: " logPath
              ))
            )
          )
        )
      )

      (princ)
    )
  )
)
