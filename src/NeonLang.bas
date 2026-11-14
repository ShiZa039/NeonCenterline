Attribute VB_Name = "NeonLang"
' =====================================================================
'  NeonCenterline — строки интерфейса (RU/EN). ФАЙЛ СГЕНЕРИРОВАН tools/gen_lang.py
'  из tools/strings.py — правьте там. Строки хранятся кодами Unicode, чтобы русский
'  текст не зависел от кодировки Windows.
' =====================================================================
Option Explicit

Private Const REG_APP As String = "NeonCenterline"

' текущий язык: ru / en (настройка UI\Lang; пусто — по языку Windows)
Public Function NeonLangCode() As String
    Dim v As String, loc As String
    v = LCase$(GetSetting(REG_APP, "UI", "Lang", ""))
    If v = "ru" Or v = "en" Then
        NeonLangCode = v
        Exit Function
    End If
    loc = ""
    On Error Resume Next
    loc = LCase$(CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\International\LocaleName"))
    On Error GoTo 0
    If Left$(loc, 2) = "ru" Or Left$(loc, 2) = "uk" Or Left$(loc, 2) = "be" Or Left$(loc, 2) = "kk" Then
        NeonLangCode = "ru"
    Else
        NeonLangCode = "en"
    End If
End Function

' строка по ключу на текущем языке; \n -> перенос строки
Public Function NeonL(ByVal key As String) As String
    Dim s As String
    If NeonLangCode() = "ru" Then
        s = LangRu(key)
    Else
        s = LangEn(key)
    End If
    NeonL = Replace(s, "\n", vbCrLf)
End Function

' то же с подстановкой {0}..{7}
Public Function NeonLF(ByVal key As String, Optional ByVal a0 As Variant = "", Optional ByVal a1 As Variant = "", _
                       Optional ByVal a2 As Variant = "", Optional ByVal a3 As Variant = "", Optional ByVal a4 As Variant = "", _
                       Optional ByVal a5 As Variant = "", Optional ByVal a6 As Variant = "", Optional ByVal a7 As Variant = "") As String
    Dim s As String
    s = NeonL(key)
    s = Replace(s, "{0}", CStr(a0)): s = Replace(s, "{1}", CStr(a1)): s = Replace(s, "{2}", CStr(a2))
    s = Replace(s, "{3}", CStr(a3)): s = Replace(s, "{4}", CStr(a4)): s = Replace(s, "{5}", CStr(a5))
    s = Replace(s, "{6}", CStr(a6)): s = Replace(s, "{7}", CStr(a7))
    NeonLF = s
End Function

' сжатая запись строки: «#» — русский режим ('?'..'~' -> А..я), «$XXXX» — знак Unicode
Private Function LD(ByVal t As String) As String
    Dim i As Long, n As Long, c As Long, cyr As Boolean, r As String
    n = Len(t)
    i = 1
    Do While i <= n
        c = AscW(Mid$(t, i, 1))
        If c = 35 Then
            cyr = Not cyr
            i = i + 1
        ElseIf c = 36 Then
            r = r & ChrW(CLng("&H" & Mid$(t, i + 1, 4)))
            i = i + 5
        ElseIf cyr And c >= 63 Then
            r = r & ChrW(&H410 + c - 63)
            i = i + 1
        Else
            r = r & Mid$(t, i, 1)
            i = i + 1
        End If
    Loop
    LD = r
End Function

Private Function LangRu(ByVal key As String) As String
    Dim s As String
    If LangRu_1(key, s) Then
    ElseIf LangRu_2(key, s) Then
    ElseIf LangRu_3(key, s) Then
    ElseIf LangRu_4(key, s) Then
    ElseIf LangRu_5(key, s) Then
    ElseIf LangRu_6(key, s) Then
    Else
        s = key
    End If
    LangRu = s
End Function

Private Function LangRu_1(ByVal key As String, s As String) As Boolean
    LangRu_1 = True
    s = ""
    Select Case key
    Case "app_title"
        s = s & LD("NeonCenterline")
    Case "center_title"
        s = s & LD("#Udlqo_j{l_~ jglg~")
    Case "mode_mono"
        s = s & LD("#Kmlm")
    Case "mode_script"
        s = s & LD("#Iropga")
    Case "script_beta"
        s = s & LD("$000A$000A$26A0 #Odegk $00ABIropga$00BB a o_fo_`mqid g o_`mq_dq ldpq_`gj{lm: l_ p")
        s = s & LD("#qzi_t `ria amfkmelz gfjmkz g jgwlgd irpig. Nomado{qd jglgg n")
        s = s & LD("#dodc odfimh.")
    Case "err_nodoc"
        s = s & LD("#Ldq mqiozqmbm cmirkdlq_.")
    Case "err_nosel"
        s = s & LD("#Azcdjgqd imlqroz `ria.")
    Case "err_noclosed"
        s = s & LD("#A azcdjdlgg ldq f_kilrqzt imlqroma.")
    Case "err_fail"
        s = s & LD("#Mwg`i_: #{0}")
    Case "err_nomodule"
        s = s & LD("#Kmcrj{ $00AB#{0}$00BB #ld l_hcdl a nomdiqd nj_bgl_. Ndodrpq_lmagqd nj_")
        s = s & LD("#bgl.")
    Case "done"
        s = s & LD("#Bmqmam (#{0}, #sodf_ #{1} #kk): m`ydiqma #{2}, #jglgh pmfc_lm #{3}.")
        s = s & LD("$000A#Aodk~ #{4} #p (o_pv$0451q #{5} #p).")
    Case "warn_thin"
        s = s & LD("#Sodf_ adjgi_ cj~ |qgt `ria: wqogt $2248 #{0} #kk, _ imluz jglgh mq")
        s = s & LD("#pqrn_}q mq io_~ l_ #{1} #kk $2014 jglgg rimomvdlz, nomn_jm udjgimk")
        s = s & LD(": {2}.$000A#Cj~ kdjimbm qdipq_ g bo_agomaig a mild ndodc o_pv$0451qmk")
        s = s & LD(" #f_c_hqd cg_kdqo sodfz = cg_kdqo bo_ado_ gjg $00ABMqpqrn imlu_ j")
        s = s & LD("#glgg mq io_~$00BB 0$20130,3 kk.")
    Case "prompt_tool"
        s = s & LD("#Cg_kdqo sodfz / o_ppdga_qdj~, kk:$000A(imldu jglgg pq_agqp~ l_ #D")
        s = s & LD("/2 #mq io_~ imlqro_)")
    Case "bad_tool"
        s = s & LD("#Ldadolzh cg_kdqo.")
    Case "settings_title"
        s = s & LD("#L_pqomhig")
    Case "lbl_tool"
        s = s & LD("#Cg_kdqo sodfz, kk:")
    Case "lbl_tool_hint"
        s = s & LD("#Pam`mclzh imldu jglgg pq_agqp~ l_ #D/2 #mq io_~ imlqro_.")
    Case "chk_ask"
        s = s & LD("#Pno_wga_q{ n_o_kdqoz nog i_ecmk f_nrpid")
    Case "lbl_lang"
        s = s & LD("#^fzi glqdosdhp_:")
    Case "lang_auto"
        s = s & LD("#?aqm (i_i a #Windows)")
    Case "lang_ru"
        s = s & LD("#Orppigh")
    Case "lang_en"
        s = s & LD("English")
    Case "chk_upd"
        s = s & LD("#Nomado~q{ m`lmajdlg~ o_f a cdl{")
    Case "lbl_version"
        s = s & LD("#Adopg~ #{0}")
    Case "btn_ok"
        s = s & LD("OK")
    Case "btn_cancel"
        s = s & LD("#Mqkdl_")
    Case "btn_update"
        s = s & LD("#Nomadogq{ m`lmajdlg~")
    Case "btn_about"
        s = s & LD("#M nombo_kkd")
    Case "lang_changed"
        s = s & LD("#^fzi `rcdq nogkdl$0451l i pmm`xdlg~k po_fr, i milr l_pqomdi $2014 no")
        s = s & LD("#g pjdcr}xdk mqiozqgg.")
    Case "upd_none"
        s = s & LD("#R a_p nmpjdcl~~ adopg~ (#{0}).")
    Case "upd_found"
        s = s & LD("#Cmpqrnl_ lma_~ adopg~ #{0} (#r a_p #{1}).$000A$000A#Vqm lmambm:$000A#{2}$000A$000A#Pi_")
        s = s & LD("#v_q{ g rpq_lmagq{#? #Lma_~ adopg~ f_o_`mq_dq nmpjd ndodf_nrpi_")
        s = s & LD(" CorelDRAW.")
    Case "upd_ready"
        s = s & LD("#M`lmajdlgd #{0} #pi_v_lm. Mlm rpq_lmagqp~ _aqmk_qgvdpig, imbc_")
        s = s & LD(" #az f_iomdqd #CorelDRAW.")
    Case "upd_pending"
        s = s & LD("#M`lmajdlgd #{0} #red pi_v_lm g rpq_lmagqp~ nmpjd f_iozqg~ #Core")
        s = s & LD("lDRAW.")
    Case "upd_err"
        s = s & LD("#Ld rc_jmp{ nomadogq{ m`lmajdlg~: #{0}")
    Case "upd_bad"
        s = s & LD("#Pi_v_llzh s_hj m`lmajdlg~ ld nmcmw$0451j: #{0}. #M`lmajdlgd mqkdld")
        s = s & LD("#lm $2014 lmar} adopg} kmelm nmpq_agq{ gf _otga_ p #GitHub (instal")
        s = s & LD("l.cmd).")
    Case "upd_why_type"
        s = s & LD("#|qm ld s_hj nj_bgl_ (pdq{ nmckdlgj_ gjg m`moa_j_ f_borfir)")
    Case "upd_why_sum"
        s = s & LD("#imlqomj{l_~ prkk_ ld pman_j_ (ec_jg #{0}$2026, #nogwjm #{1}$2026)")
    Case "upd_why_nohash"
        s = s & LD("#ld rc_jmp{ nmpvgq_q{ imlqomj{lr} prkkr (#certutil #g #PowerShel")
        s = s & LD("l #ldcmpqrnlz)")
    Case "upd_nourl"
        s = s & LD("#?codp m`lmajdlgh ld l_pqomdl.")
    Case "upd_nogms"
        s = s & LD("#Ld l_hcdl s_hj nj_bgl_ (.#gms) #cj~ f_kdlz.")
    Case Else
        LangRu_1 = False
    End Select
End Function

Private Function LangRu_2(ByVal key As String, s As String) As Boolean
    LangRu_2 = True
    s = ""
    Select Case key
    Case "about"
        s = s & LD("NeonCenterline {0}$000A$000A#Udlqo_j{l_~ jglg~ nm imlqro_k `ria cj~ s")
        s = s & LD("#odfdomaig i_l_j_ nmc bg`igh ldml. Ilmnig: $00ABKmlm$00BB $2014 cj~ kmlmw")
        s = s & LD("#gogllzt (ldmlmazt) wogsqma, $00ABIropga$00BB $2014 cj~ orimngplzt.$000A$000A$00A9 20")
        s = s & LD("26 SAB:CORE. #@dpnj_qlm cj~ j}`mbm gpnmj{fma_lg~, a qmk vgpjd")
        s = s & LD(" #imkkdovdpimbm. Nomc_e_ g ndodn_imai_ nj_bgl_ f_nodxdlz (pk.")
        s = s & LD(" LICENSE).")
    Case "about_donate"
        s = s & LD("$000A$000A#Nj_bgl `dpnj_qlzh. Dpjg ml a_k nmkmb_dq $2014 kmelm nm`j_bmc_o")
        s = s & LD("#gq{ _aqmo_. Mqiozq{ pqo_lgur#?")
    Case "addon_done"
        s = s & LD("#N_ldj{ glpqorkdlqma pimngoma_l_ a$000A#{0}$000A$000A#Nmjlmpq{} ndodf_nrpqg")
        s = s & LD("#qd #CorelDRAW: #n_ldj{ $00ABLdml$00BB ndodpqomgqp~ p_k_ nog f_nrpid, n")
        s = s & LD("#og mqiozqgg cmirkdlq_ gjg nog ndoamh imk_lcd nj_bgl_. Dpjg n")
        s = s & LD("#_ldjg ld agclm: Milm > N_ldjg glpqorkdlqma > $00ABLdml (#NeonCent")
        s = s & LD("erline)$00BB.")
    Case "addon_fail"
        s = s & LD("#S_hjz n_ldjg ld pimngoma_jgp{ a$000A#{0}$000A$000A{1}$000A$000A#F_nrpqgqd k_iomp #N")
        s = s & LD("eonCenterline.NeonAddon.InstallToolbar #dx$0451 o_f g nmcqadocgqd")
        s = s & LD(" #f_nomp no_a _ckglgpqo_qmo_ gjg f_iomhqd #CorelDRAW #g pimngor")
        s = s & LD("#hqd qrc_ aorvlr} n_nir #addon\NeonCenterline #gf _otga_ nj_bgl")
        s = s & LD("#_.")
    Case "addon_nolog"
        s = s & LD("#Imngoma_lgd ld f_nrpqgjmp{: amfkmelm, f_nomp no_a _ckglgpqo_")
        s = s & LD("#qmo_ `zj mqijml$0451l.")
    Case "err_nosel_any"
        s = s & LD("#Lgvdbm ld azcdjdlm.")
    Case "len_title"
        s = s & LD("#Cjgl_ imlqroma g jglgh")
    Case "len_contours"
        s = s & LD("#Imlqroz (odfi_): m`ydiqma #{0}, #imlqroma #{1}$000A#M`x_~ cjgl_: #{2}")
    Case "len_neon"
        s = s & LD("#Udlqo_j{lzd jglgg (ldml): irpima #{0}$000A#M`x_~ cjgl_: #{1}$000A#P_kzh ")
        s = s & LD("#cjgllzh irpmi: #{2}")
    Case "len_none"
        s = s & LD("#A azcdjdlgg ldq iogazt, cjglr imqmozt kmelm nmpvgq_q{.")
    Case "rep_undo"
        s = s & LD("#MI $2014 mpq_agq{. Mqkdl_ $2014 mqkdlgq{ ap$0451 pcdj_llmd (i_i #Ctrl+Z).")
    Case "rep_undo_fail"
        s = s & LD("#Mqkdlgq{ ld nmjrvgjmp{ $2014 l_ekgqd #Ctrl+Z.")
    Case "rep_kept"
        s = s & LD("#Mogbgl_jz mpq_ajdlz: odfrj{q_q $2014 imng~kg nmadot lgt.")
    Case "chk_keep"
        s = s & LD("#Mpq_agq{ mogbgl_j (odfrj{q_q $2014 imngdh nmadot)")
    Case "lbl_keep_orig"
        s = s & LD("#Mpq_aj~q{ mogbgl_j:")
    Case "chk_keep_clean"
        s = s & LD("#mvgpqi_")
    Case "chk_keep_split"
        s = s & LD("#o_f`gai_")
    Case "chk_keep_simp"
        s = s & LD("#rnomxdlgd")
    Case "prog_start"
        s = s & LD("#Nmcbmqmai_$2026")
    Case "prog_step"
        s = s & LD("{0}: {1} #gf #{2}")
    Case "prog_time"
        s = s & LD("#Gc$0451q o_`mq_: #{0} #p")
    Case "prog_read"
        s = s & LD("#Vqdlgd imlqroma")
    Case "prog_core"
        s = s & LD("#Udlqo_j{lzd jglgg, bornnz `ria")
    Case "prog_emit"
        s = s & LD("#Nmpqomdlgd jglgh")
    Case "prog_objects"
        s = s & LD("#M`ydiqz")
    Case "prog_compare"
        s = s & LD("#Po_aldlgd smok")
    Case "prog_nest_geo"
        s = s & LD("#Imlqroz cdq_jdh (#{0})")
    Case "prog_nest_pack"
        s = s & LD("#O_pij_ci_, a_og_lq #{0} #gf #{1}")
    Case "prog_nest_place"
        s = s & LD("#Cdq_jg l_ jgpqz (#{0})")
    Case "prog_count"
        s = s & LD("{0}: {1}")
    Case "prog_measure"
        s = s & LD("#Gfkdodlgd")
    Case "prog_find"
        s = s & LD("#Nmgpi udlqo_j{lzt jglgh")
    Case "prog_cuts"
        s = s & LD("#Odfz")
    Case "prog_collect"
        s = s & LD("#P`mo m`ydiqma")
    Case "prog_groups"
        s = s & LD("#Bornnz")
    Case "prog_cutfit"
        s = s & LD("#Nmcbmli_ nmc io_qlmpq{")
    Case "sel_none"
        s = s & LD("#L_ pqo_lgud ldq udlqo_j{lzt jglgh nj_bgl_.")
    Case "unit_mm"
        s = s & LD("#kk")
    Case "unit_m"
        s = s & LD("#k")
    Case "cut_title"
        s = s & LD("#Io_qlmpq{ odf_")
    Case Else
        LangRu_2 = False
    End Select
End Function

Private Function LangRu_3(ByVal key As String, s As String) As Boolean
    LangRu_3 = True
    s = ""
    Select Case key
    Case "cut_nostep"
        s = s & LD("#Io_qlmpq{ odf_ ld f_c_l_ (L_pqomhig).")
    Case "cut_report"
        s = s & LD("#Nmcbmli_ nmc odf #{0} #kk: rimomvdlm irpima #{1} (#podf_lm apdbm")
        s = s & LD(" {2} #kk), red io_qlz #{3}, #ld nmcmbl_lm #{4} (#f_kilrqzd, `df p")
        s = s & LD("#am`mclzt imluma gjg imomvd w_b_).")
    Case "cut_sum_classic"
        s = s & LD("#Ldml io_qlm odfr #{0} #kk: #{1} #kk, mpq_qig (ld f_nmjlgqp~ ldml")
        s = s & LD("#mk) #{2} #kk")
    Case "cut_sum_neon2"
        s = s & LD("#Jdlq_ io_qlm odfr #{0} #kk: #{1} #kk, n_f `df jdlqz #{2} #kk")
    Case "lbl_type"
        s = s & LD("#Qgn ldml_:")
    Case "type_classic"
        s = s & LD("#Padqmcgmclzh ldml (jdlq_ alrqog)")
    Case "type_neon2"
        s = s & LD("#Ldml 2.0 (o_ppdga_qdj{ a n_fr, jdlq_ mqcdj{lm)")
    Case "lbl_cut"
        s = s & LD("#Io_qlmpq{ odf_, kk:")
    Case "chk_fit"
        s = s & LD("#Nmcbml~q{ jglgg nmc io_qlmpq{ odf_ (rimo_vga_q{ pam`mclzd im")
        s = s & LD("#luz)")
    Case "lbl_trim"
        s = s & LD("#Rimo_vga_q{:")
    Case "trim_both"
        s = s & LD("#nmomalr p m`mgt imluma")
    Case "trim_end"
        s = s & LD("#qmj{im p imlu_ jglgg")
    Case "chk_ask_run"
        s = s & LD("#Nmi_fza_q{ |qm milm nog i_ecmk f_nrpid (azij}vgq{ $2014 a L_pqom")
        s = s & LD("#hi_t)")
    Case "btn_run"
        s = s & LD("#Nmpqomgq{")
    Case "upd_net_hint"
        s = s & LD("$000A$000A#Amfkmelm, #GitHub #ldcmpqrndl gf a_wdh pdqg (r v_pqg noma_hc")
        s = s & LD("#doma ml `jmigordqp~). Kmelm pi_v_q{ lmar} adopg} aorvlr}: #gi")
        s = s & LD("thub.com/ShiZa039/NeonCenterline/releases")
    Case "lbl_ref"
        s = s & LD("#Iropga $2014 wogsq i_i m`o_fdu (`df rv$0451q_ sodfz):")
    Case "ref_off"
        s = s & LD("#ldq: sodfdoma_q{ p_k imlqro")
    Case "ref_tip"
        s = s & LD("#c_, jglg~ cm imlvgi_ wqogt_")
    Case "ref_half"
        s = s & LD("#c_, mqpqrn l_ nmjwgoglz wqogt_")
    Case "lbl_endoff"
        s = s & LD("#Mqpqrn imlu_ jglgg mq io_~, kk:")
    Case "endoff_default"
        s = s & LD("D/2 (#o_cgrp sodfz)")
    Case "fallback_lang"
        s = s & LD("#^fzi glqdosdhp_: C_ $2014 orppigh, Ldq $2014 #English, #Mqkdl_ $2014 _aqm.")
    Case "fallback_upd"
        s = s & LD("#Nomado~q{ m`lmajdlg~ o_f a cdl{#?")
    Case "fallback_ask"
        s = s & LD("#Pno_wga_q{ cg_kdqo nog i_ecmk f_nrpid#?")
    Case "clean_title"
        s = s & LD("#Mvgpqgq{ imlqro")
    Case "clean_none"
        s = s & LD("#A azcdjdlgg ldq iogazt g qdipq_ (no~kmrbmj{lgig g |jjgnpz ld")
        s = s & LD(" #mvgx_}qp~ $2014 l_jmedlgh a lgt ldq).")
    Case "clean_nochange"
        s = s & LD("#L_jmedlgh g ndqdj{ ld l_hcdlm $2014 m`ydiqz ld gfkdldlz.")
    Case "clean_report"
        s = s & LD("#Mvgxdlm m`ydiqma: #{0} #gf #{1}.$000A#Ndodpdvdlgh imlqroma r`o_lm: #{")
        s = s & LD("2}.$000A#F_kilrqzt imlqroma `zjm #{3}, #pq_jm #{4}.$000A$000A#I_eczh m`ydiq $2014")
        s = s & LD(" #mcl_ ioga_~.")
    Case "clean_skipped"
        s = s & LD("#Nomnrxdlm m`ydiqma (ld iogazd g ld qdipq): #{0}.")
    Case "simp_title"
        s = s & LD("#Rnompqgq{ imlqro")
    Case "simp_none"
        s = s & LD("#A azcdjdlgg ldq iogazt. Qdipq g sgbroz pl_v_j_ ndodadcgqd a ")
        s = s & LD("#iogazd (#Ctrl+Q).")
    Case "simp_report"
        s = s & LD("#Cmnrpi #{0} #kk: rnomxdlm iogazt #{1} #gf #{2}.$000A#Rfjma `zjm #{3}, #p")
        s = s & LD("#q_jm #{4} (#kdl{wd l_ #{5}%).")
    Case "simp_skipped"
        s = s & LD("#Nomnrxdlm m`ydiqma, imqmozd ld iogazd: #{0} (#ndodadcgqd a iog")
        s = s & LD("#azd $2014 #Ctrl+Q).")
    Case "lbl_simp"
        s = s & LD("#Rnomxdlgd imlqro_, cmnrpi, kk:")
    Case "prompt_simp"
        s = s & LD("#Rnomxdlgd imlqro_: cmnrpi, kk$000A(l_pimj{im ioga_~ kmedq mqmhqg")
        s = s & LD(" #mq gptmclmh; 0.1 $2014 ldf_kdqlm)")
    Case "bad_simp"
        s = s & LD("#Cmnrpi rnomxdlg~ $2014 vgpjm mq 0.01 cm 5 kk.")
    Case "lbl_simp_tol"
        s = s & LD("#Cmnrpi, kk:")
    Case "simp_hint"
        s = s & LD("#L_pimj{im ioga_~ kmedq mqmhqg mq gptmclmh. 0.1 kk $2014 ldf_kdql")
        s = s & LD("#m l_ bj_f g cj~ sodfz; `mj{wd $2014 pgj{ldd pbj_ega_dq $00ABoa_lzd$00BB ")
        s = s & LD("#imlqroz nmpjd qo_ppgomaig.")
    Case "btn_simp"
        s = s & LD("#Rnompqgq{")
    Case "chk_simp_ask"
        s = s & LD("#Nmi_fza_q{ milm rnomxdlg~ nog i_ecmk f_nrpid")
    Case Else
        LangRu_3 = False
    End Select
End Function

Private Function LangRu_4(ByVal key As String, s As String) As Boolean
    LangRu_4 = True
    s = ""
    Select Case key
    Case "tb_outdated"
        s = s & LD("#N_ldj{ $00ABLdml$00BB a #CorelDRAW #rpq_odj_: a |qmh adopgg nj_bgl_ l_")
        s = s & LD(" #ldh dpq{ lmazd ilmnig. M`lmagq{ n_ldj{ pdhv_p#?$000A$000AWindows #nmn")
        s = s & LD("#ompgq no_a_ _ckglgpqo_qmo_. F_qdk nmjlmpq{} ndodf_nrpqgqd #Co")
        s = s & LD("relDRAW $2014 #lmazd ilmnig nm~a~qp~ p_kg.")
    Case "tb_reset_done"
        s = s & LD("#N_ldj{ $00ABLdml$00BB p`omwdl_ i l_`mor ilmnmi |qmh adopgg nj_bgl_.$000A")
        s = s & LD("$000A{0}")
    Case "tb_reset_none"
        s = s & LD("#N_ldj{ $00ABLdml$00BB ld rc_jmp{ p`ompgq{.$000A$000A#{0}$000A#Dpjg n_ldjg ldq a pn")
        s = s & LD("#gpid (Milm > N_ldjg glpqorkdlqma), rpq_lmagqd d$0451: k_iomp #Neo")
        s = s & LD("nCenterline.NeonAddon.InstallToolbar, #f_qdk nmjlmpq{} ndodf_")
        s = s & LD("#nrpqgqd #CorelDRAW.")
    Case "tb_files_old"
        s = s & LD("#S_hjz n_ldjg $00ABLdml$00BB a n_nid #Addons CorelDRAW $2014 #mq pq_omh ado")
        s = s & LD("#pgg nj_bgl_ (gjg gt ldq):$000A#{0}$000A#Lredl l_`mo ilmnmi #[NeonToolba")
        s = s & LD("r {1}].$000A$000A#M`lmagq{ s_hjz n_ldjg pdhv_p#? Windows #nmnompgq no_a")
        s = s & LD("#_ _ckglgpqo_qmo_, f_qdk nmjlmpq{} ndodf_nrpqgqd #CorelDRAW.")
    Case "grom_title"
        s = s & LD("#J}adopz")
    Case "grom_nosel"
        s = s & LD("#Azcdjgqd `_lldo (nmcjmeir): no~kmrbmj{lgi gjg f_kilrqzh imlq")
        s = s & LD("#ro.")
    Case "grom_none"
        s = s & LD("#A azcdjdlgg ldq f_kilrqmbm imlqro_, a imqmozh j}adopz nmkdx_")
        s = s & LD("#}qp~ p q_igk mqpqrnmk mq io_~.")
    Case "grom_report"
        s = s & LD("#J}adopma: #{0} (#m`ydiqma: #{1}).$000A#Mq io_~ cm udlqo_ j}adop_ #{2}")
        s = s & LD(" #kk, o_ppqm~lgd kdecr j}adop_kg ld `mj{wd #{3} #kk, qmvi_ $00D8#{4}")
        s = s & LD(" #kk.$000AQmvig $2014 l_ pjmd $00ABJ}adopz$00BB, nm bornnd l_ m`ydiq.")
    Case "grom_report_reinf"
        s = s & LD("#Rpgjdlgd rbjma: dx$0451 nm qmvid a #{0} #kk mq i_ecmbm rbj_.")
    Case "grom_skipped"
        s = s & LD("#Nomnrxdlm m`ydiqma: #{0} (#ldq f_kilrqmbm imlqro_ gjg mqpqrn `")
        s = s & LD("#mj{wd o_fkdo_).")
    Case "grom_why_nocurve"
        s = s & LD("#ld rc_jmp{ nmjrvgq{ imlqro")
    Case "grom_why_open"
        s = s & LD("#ldq f_kilrqmbm imlqro_ (nmcnrqdh: #{0}, #f_kilrqzt: #{1}; #imluz")
        s = s & LD(" #ld ptmc~qp~)")
    Case "grom_why_small"
        s = s & LD("#j}adopz ld nmkdx_}qp~ p q_igk mqpqrnmk mq io_~")
    Case "grom_t_curve"
        s = s & LD("#ioga_~")
    Case "grom_t_rect"
        s = s & LD("#no~kmrbmj{lgi")
    Case "grom_t_ellipse"
        s = s & LD("#|jjgnp")
    Case "grom_t_poly"
        s = s & LD("#klmbmrbmj{lgi")
    Case "grom_t_text"
        s = s & LD("#qdipq")
    Case "grom_t_other"
        s = s & LD("#m`ydiq (qgn #{0})")
    Case "grom_inner"
        s = s & LD("#Alrqog `_lldo_ (ogprlmi, qdipq, o_kig) $2014 `df j}adopma: #{0}.")
    Case "grom_layer"
        s = s & LD("#J}adopz")
    Case "grom_dot"
        s = s & LD("#J}adop")
    Case "grom_group"
        s = s & LD("#J}adopz: #{0}")
    Case "lbl_grom_edge"
        s = s & LD("#Mq io_~ cm udlqo_ j}adop_, kk:")
    Case "lbl_grom_step"
        s = s & LD("#O_ppqm~lgd kdecr j}adop_kg, ld `mj{wd, kk:")
    Case "lbl_grom_dot"
        s = s & LD("#Cg_kdqo qmvig, kk:")
    Case "lbl_grom_color"
        s = s & LD("#Uadq qmvig:")
    Case "grom_c_black"
        s = s & LD("#V$0451olzh (#K100)")
    Case "grom_c_red"
        s = s & LD("#Io_plzh (#M100 Y100)")
    Case "grom_c_magenta"
        s = s & LD("#Nronrolzh (#M100)")
    Case "grom_c_blue"
        s = s & LD("#Pglgh (#C100 M100)")
    Case "grom_c_green"
        s = s & LD("#Fdj$0451lzh (#C100 Y100)")
    Case "grom_c_white"
        s = s & LD("#@djzh")
    Case "grom_c_other"
        s = s & LD("#Corbmh$2026")
    Case "chk_grom_reinf"
        s = s & LD("#Rpgjdlgd rbjma: dx$0451 qmvi_ a, kk")
    Case "chk_grom_hem"
        s = s & LD("#Imlqro nmc nmcbg`, aznrpi, kk:")
    Case "chk_grom_layer"
        s = s & LD("#J}adopz $2014 l_ mqcdj{lzh pjmh $00ABJ}adopz$00BB (gl_vd $2014 l_ pjmh `_lld")
        s = s & LD("#o_)")
    Case "bad_grom_hem"
        s = s & LD("#Aznrpi nmc nmcbg` $2014 mq 0 cm 1000 kk.")
    Case "grom_hem_name"
        s = s & LD("#Nmcbg` #{0} #kk")
    Case "grom_hem_rep"
        s = s & LD("#Imlqro nmc nmcbg`: aznrpi #{0} #kk l_orer mq io_~ (m`ydiqma: #{")
        s = s & LD("1}) $2014 #l_ pjmd `_lldo_.")
    Case Else
        LangRu_4 = False
    End Select
End Function

Private Function LangRu_5(ByVal key As String, s As String) As Boolean
    LangRu_5 = True
    s = ""
    Select Case key
    Case "grom_hint"
        s = s & LD("#A i_ecmk rbjr $2014 j}adop, l_ pqmoml_t $2014 o_almkdolm, ld oded f_")
        s = s & LD("#c_llmbm o_ppqm~lg~. Piorbj$0451llzh rbmj $2014 qmed rbmj; r iorb_ g ")
        s = s & LD("#ma_j_ $2014 o_almkdolm nm imlqror. Imlqro nmc nmcbg` $2014 jglg~ odf")
        s = s & LD("#_ l_ aznrpi l_orer mq io_~ `_lldo_ (j}adopz $2014 mq io_~, i_i `")
        s = s & LD("#df nmcbg`_).")
    Case "btn_grom"
        s = s & LD("#O_ppq_agq{")
    Case "bad_grom"
        s = s & LD("#Nomado{qd fl_vdlg~: mq io_~ 0$2013500 kk, o_ppqm~lgd 20$201310000 kk")
        s = s & LD(", #qmvi_ 0$2013100 kk.")
    Case "split_title"
        s = s & LD("#O_f`gq{ l_ `riaz")
    Case "split_none"
        s = s & LD("#A azcdjdlgg ldq iogazt gjg qdipq_ gf ldpimj{igt `ria $2014 o_f`g")
        s = s & LD("#a_q{ ldvdbm.")
    Case "split_report"
        s = s & LD("#O_f`gqm m`ydiqma: #{0}, #nmjrvgjmp{ mqcdj{lzt `ria g fl_ima: #{")
        s = s & LD("1}. #Czoig mpq_jgp{ a pamgt `ria_t, qmvig g _iudlqz ($0401, H, #i,")
        s = s & LD(" !, ?) $2014 #pm pamdh `riamh; cdq_jg alrqog czomi (`ria_ a $00AB$00AE$00BB, ")
        s = s & LD("#qmvi_ a lrjd) $2014 mqcdj{lzkg m`ydiq_kg. M`ydiqma gf mclmh `ria")
        s = s & LD("#z $2014 ld qomlrqz: #{2}.$000A$000A#Dpjg `riaz l_jdf_}q corb l_ corb_, pl_")
        s = s & LD("#v_j_ $00ABMvgpqgq{ imlqro$00BB.")
    Case "noodle_title"
        s = s & LD("#J_nw_ (nmjmpz)")
    Case "lbl_noodle_len"
        s = s & LD("#Cjgl_ j_nwg, kk:")
    Case "lbl_noodle_wid"
        s = s & LD("#Wgogl_ j_nwg, kk:")
    Case "lbl_noodle_cnt"
        s = s & LD("#Imjgvdpqam, wq:")
    Case "lbl_noodle_overh"
        s = s & LD("#Tampqz bmogfmlq_j{lzt odfma, kk:")
    Case "lbl_noodle_overv"
        s = s & LD("#Tampqz adoqgi_j{lzt mqodfma, kk:")
    Case "chk_noodle_tool"
        s = s & LD("#Sodf_ $2014 cg_kdqo, kk:")
    Case "bad_noodle_tool"
        s = s & LD("#Cg_kdqo sodfz $2014 mq 0 cm 100 kk.")
    Case "noodle_hint"
        s = s & LD("#Nmjmpz cj~ njmqqdo_, j_fdo_, sodfdo_. Nmo~cmi odfig $2014 nmo~cm")
        s = s & LD("#i m`ydiqma: bmogfmlq_j{lzd odfz padotr algf fgbf_bmk (l_no_a")
        s = s & LD("#jdlg~ vdodcr}qp~, pq_lmi ld tmcgq atmjmpqr}), a p_kmk imlud ")
        s = s & LD("#pngpi_ $2014 ca_ adoqgi_j{lzt mqodf_, qmed nm tmcr. P b_jmvimh $00AB")
        s = s & LD("#Sodf_$00BB odfz $2014 nm udlqor sodfz: bmogfmlq_j{lzd gcrq vdodf wgo")
        s = s & LD("#glr + cg_kdqo, mqodfz $2014 l_ o_cgrp f_ imlu_kg nmjmp, g nmjmpz")
        s = s & LD(" #nmpjd sodfz aztmc~q omalm f_c_llmbm o_fkdo_. Tampqz $2014 aztmc")
        s = s & LD(" #odf_ f_ jglg} corbmbm odf_, vqm`z nmjmp_ mqcdjgj_p{ udjgimk")
        s = s & LD(": #j_fdor g njmqqdor ta_q_dq 1$20132 kk, sodfd $2014 2$20135 kk f_n_p_. A")
        s = s & LD("#p$0451 $2014 mclmh bornnmh a udlqod pqo_lguz.")
    Case "btn_noodle"
        s = s & LD("#L_odf_q{")
    Case "bad_noodle"
        s = s & LD("#Nomado{qd fl_vdlg~: cjgl_ 0$2013100000 kk, wgogl_ 0$201310000 kk, im")
        s = s & LD("#jgvdpqam $2014 udjmd 1$20131000, tampqz 0$20131000 kk.")
    Case "noodle_group"
        s = s & LD("#J_nw_: #{0} #wq $00B7 #{1} $00D7 {2} #kk")
    Case "noodle_report"
        s = s & LD("#J_nw_: #{0} #nmjmp #{1} $00D7 {2} #kk.$000AOdfz: #{3} #bmogfmlq_j{lzt pado")
        s = s & LD("#tr algf fgbf_bmk, f_qdk 2 adoqgi_j{lzt mqodf_ $2014 a imlud pngp")
        s = s & LD("#i_; tampqz: bmogfmlq_j{lzt #{4} #kk, adoqgi_j{lzt #{5} #kk. B_`_")
        s = s & LD("#ogq p tampq_kg: #{6} $00D7 {7} #kk.")
    Case "noodle_rep_tool"
        s = s & LD("#Sodf_ $2300#{0} #kk: odfz $2014 nm udlqor sodfz, bmogfmlq_j{lzd vdodf ")
        s = s & LD("{1} #kk, mqodfz vdodf #{2} #kk $2014 nmjmpz nmpjd sodfz #{3} $00D7 {4} #k")
        s = s & LD("#k.")
    Case "noodle_rep_tail"
        s = s & LD("#Mclmh bornnmh a udlqod pqo_lguz; nmo~cmi m`ydiqma $2014 nmo~cmi ")
        s = s & LD("#odfig.")
    Case "same_title"
        s = s & LD("#Mcgl_imazd $2014 a bornnz")
    Case "same_nosel"
        s = s & LD("#Azcdjgqd m`ydiqz: iogazd, qdipq, sgbroz.")
    Case "same_none"
        s = s & LD("#Mcgl_imazt m`ydiqma ld l_hcdlm: apd #{0} #o_flzd.$000A$000AMcgl_imazd ")
        s = s & LD("$2014 #mcl_ smok_ p qmvlmpq{} cm 0,1 kk (`df nmamomq_ g fdoi_j_) ")
        s = s & LD("#g mcgl uadq.")
    Case "same_report"
        s = s & LD("#M`ydiqma: #{0}, #o_flzt smok: #{1}.$000A#Mcgl_imazd pm`o_lz a bornnz")
        s = s & LD(": #bornn #{2}, #a lgt m`ydiqma #{3}. #@df n_oz: #{4}.$000A$000A#Mcgl_imazd ")
        s = s & LD("$2014 #mcl_ smok_ p qmvlmpq{} cm 0,1 kk (`df nmamomq_ g fdoi_j_) ")
        s = s & LD("#g mcgl uadq.")
    Case "same_list"
        s = s & LD("#Smokz (wqri $2014 b_`_ogq $2014 uadq):")
    Case "same_line"
        s = s & LD("{0} #wq $2014 #{1} $00D7 {2} #kk $2014 #{3}")
    Case "same_more"
        s = s & LD("$2026 #g dx$0451 smok: #{0}")
    Case "same_group"
        s = s & LD("#Mcgl_imazd: #{0} #wq $00B7 #{1} $00D7 {2} #kk")
    Case "same_failed"
        s = s & LD("#Ld rc_jmp{ pbornngoma_q{ l_`moma: #{0} (#m`ydiqz l_ o_flzt pjm")
        s = s & LD("#~t#?).")
    Case "no_color"
        s = s & LD("#`df uadq_")
    Case "nest_title"
        s = s & LD("#O_pij_ci_ l_ jgpqz (`dq_)")
    Case "nest_beta"
        s = s & LD("$26A0 #O_pij_ci_ a o_fo_`mqid: nm imlqror $2014 l_ pdqid mimjm 1 kk (")
        s = s & LD("#l_ `mj{wgt jgpq_t $2014 iornldd), cdq_jg kmbrq apq_q{ a czoig `r")
        s = s & LD("#ia g alrqo{ corbgt cdq_jdh (b_jmvi_ $00ABk_qo$0451wimh$00BB a mild). Iml")
        s = s & LD("#qro cdq_jg cmjedl `zq{ f_kilrqzk $2014 gl_vd alrqo{ ld$0451 kmedq ap")
        s = s & LD("#q_q{ cdq_j{ c_ed `df $00ABk_qo$0451wig$00BB. Nomado~hqd o_pij_cir ndodc ")
        s = s & LD("#odfimh.")
    Case "nest_nosel"
        s = s & LD("#Azcdjgqd cdq_jg cj~ o_pij_cig.")
    Case "nest_layer"
        s = s & LD("#O_pij_ci_")
    Case "nest_frame"
        s = s & LD("#O_pij_ci_: o_ki_ jgpq_ #{0}")
    Case "nest_labname"
        s = s & LD("#O_pij_ci_: nmcngp{ jgpq_ #{0}")
    Case "nest_group"
        s = s & LD("#O_pij_ci_: jgpq #{0} #gf #{1} $00B7 {2}")
    Case "nest_label"
        s = s & LD("#Jgpq #{0} #gf #{1} $00B7 {2} $00D7 {3} #kk $00B7 #{4} $00B7 #cdq_jdh #{5} $00B7 #f_nmjld")
        s = s & LD("#lgd #{6}%")
    Case "nest_report"
        s = s & LD("#O_fjmedlm cdq_jdh: #{0}, #jgpqma: #{1}.$000A{2}#Nmjd mq io_~ #{3} #kk,")
        s = s & LD(" #f_fmo kdecr cdq_j~kg #{4} #kk, #{5}.")
    Case Else
        LangRu_5 = False
    End Select
End Function

Private Function LangRu_6(ByVal key As String, s As String) As Boolean
    LangRu_6 = True
    s = ""
    Select Case key
    Case "nest_how_c0"
        s = s & LD("#nm imlqror, `df nmamomq_")
    Case "nest_how_c1"
        s = s & LD("#nm imlqror, p nmamomqmk l_ 180$00B0")
    Case "nest_how_c2"
        s = s & LD("#nm imlqror, p nmamomq_kg nm 90$00B0")
    Case "nest_how_box"
        s = s & LD("#nm b_`_ogq_k, `df nmamomq_")
    Case "nest_how_noin"
        s = s & LD(", #ld alrqo{ corbgt cdq_jdh")
    Case "nest_rep_sheet"
        s = s & LD("#Jgpq #{0} $00D7 {1} #kk (cjgl_ $00D7 wgogl_). ")
    Case "nest_rep_auto"
        s = s & LD("#Wgogl_ jgpq_ #{0} #kk, cjgl_ nm o_pij_cid (#{1}$2013{2} #kk, aadot c")
        s = s & LD("#m 10 kk). ")
    Case "nest_rep_color"
        s = s & LD("{0}: #jgpqma #{1}, #cdq_jdh #{2}, #f_nmjldlgd #{3}%")
    Case "nest_rep_widths"
        s = s & LD(" (#cjgl_: #{0})")
    Case "nest_rep_tail"
        s = s & LD("#I_eczh jgpq $2014 bornn_ pno_a_ mq azcdjdlg~ l_ pjmd cdq_jdh: o_")
        s = s & LD("#ki_, nmcngp{ g cdq_jg (nmcngpg kmelm mqij}vgq{ a mild). F_nm")
        s = s & LD("#jldlgd $2014 nm b_`_ogq_k cdq_jdh.")
    Case "nest_rep_tail_c"
        s = s & LD("#I_eczh jgpq $2014 bornn_ pno_a_ mq azcdjdlg~ l_ pjmd cdq_jdh: o_")
        s = s & LD("#ki_, nmcngp{ g cdq_jg (nmcngpg kmelm mqij}vgq{ a mild). F_nm")
        s = s & LD("#jldlgd $2014 nm njmx_cg cdq_jdh (czoig ld a pv$0451q).")
    Case "nest_unfit"
        s = s & LD("#@mj{wd o_`mvdbm nmj~ jgpq_, mpq_ajdlz l_ kdpqd: #{0}.")
    Case "nest_none"
        s = s & LD("#Lg mcl_ cdq_j{ ld nmkdx_dqp~ l_ jgpq p q_igkg nmj~kg.")
    Case "all_colors"
        s = s & LD("#apd uadq_")
    Case "lbl_nest_size"
        s = s & LD("#Bmqmazh o_fkdo:")
    Case "lbl_nest_len"
        s = s & LD("#Cjgl_ jgpq_, kk:")
    Case "lbl_nest_wid"
        s = s & LD("#Wgogl_ jgpq_, kk:")
    Case "prompt_nest_size"
        s = s & LD("#Jgpq: cjgl_ $00D7 wgogl_, kk (l_nogkdo 900#x600):")
    Case "lbl_nest_margin"
        s = s & LD("#Nmjd mq io_~ jgpq_, kk:")
    Case "lbl_nest_gap"
        s = s & LD("#F_fmo kdecr cdq_j~kg, kk:")
    Case "chk_nest_color"
        s = s & LD("#I_eczh uadq $2014 l_ pamg jgpqz")
    Case "lbl_nest_method"
        s = s & LD("#O_pij_ci_:")
    Case "nest_m_contour"
        s = s & LD("#nm imlqror $2014 njmqldd")
    Case "nest_m_rot180"
        s = s & LD("#nm imlqror, nmamomq l_ 180$00B0")
    Case "nest_m_rot90"
        s = s & LD("#nm imlqror, nmamomqz nm 90$00B0")
    Case "nest_m_box"
        s = s & LD("#nm b_`_ogq_k $2014 `zpqodd")
    Case "chk_nest_groups"
        s = s & LD("#Bornnz $2014 udjgimk (i_ec_~ bornn_ $2014 mcl_ cdq_j{)")
    Case "chk_nest_keep"
        s = s & LD("#Mpq_agq{ mogbgl_j $2014 l_ jgpqz apq_}q imngg cdq_jdh")
    Case "nest_rep_kept"
        s = s & LD("#Mogbgl_jz mpq_ajdlz l_ kdpqd, l_ jgpq_t $2014 imngg.")
    Case "chk_nest_inside"
        s = s & LD("#Cdq_jg alrqo{ corbgt (k_qo$0451wimh): a czoig `ria, o_kig")
    Case "lbl_nest_fill"
        s = s & LD("#F_nmjl~q{ jgpq:")
    Case "nest_fill_cols"
        s = s & LD("#pqmj`u_kg $2014 padotr algf")
    Case "nest_fill_rows"
        s = s & LD("#o~c_kg $2014 pjda_ l_no_am")
    Case "nest_fill_auto"
        s = s & LD("#i_i njmqldd (_aqm)")
    Case "chk_nest_labels"
        s = s & LD("#Nmcngpg jgpqma (l_c o_kimh, a bornnd jgpq_)")
    Case "chk_nest_auto"
        s = s & LD("#Cjgl_ nm o_pij_cid, ld kdl{wd, kk:")
    Case "nest_hint"
        s = s & LD("#@dq_. Cjgl_ $2014 nm bmogfmlq_jg, wgogl_ $2014 nm adoqgi_jg; o_fkdo ")
        s = s & LD("#j}`mh: angwgqd pamh (pqmj j_fdo_, f_bmqmai_), ml nm~agqp~ a ")
        s = s & LD("#bmqmazt. Nm imlqror $2014 njmqldd: cdq_jg atmc~q corb a corb_, $00AB")
        s = s & LD("#k_qo$0451wimh$00BB $2014 dx$0451 g a czoig `ria g o_kig; nmamomq $2014 dpjg k_qd")
        s = s & LD("#og_jr ap$0451 o_alm (`df qdipqroz g l_no_ajdlg~). Nm b_`_ogq_k $2014")
        s = s & LD(" #`zpqodd. Lg mcl_ cdq_j{ ld `jged nmj~ i io_} g f_fmo_ i pmp")
        s = s & LD("#dcldh; cj~ sodfz f_fmo $2014 ld kdl{wd d$0451 cg_kdqo_. P b_jmvimh $00AB")
        s = s & LD("#Cjgl_ nm o_pij_cid$00BB jgpq m`odf_dqp~ nm cdq_j~k. Bornnz $2014 udj")
        s = s & LD("#gimk (q_`jgvig, p`moig); `df b_jmvig o_f`go_}qp~ l_ |jdkdlqz")
        s = s & LD(" $2014 #q_i rcm`lm cj~ gknmoq_ #SVG/DXF #mclmh bornnmh. $00ABMcgl_imazd")
        s = s & LD("$00BB $2014 #apdbc_ nmwqrvlm.")
    Case "btn_nest"
        s = s & LD("#O_fjmegq{")
    Case "bad_nest"
        s = s & LD("#Nomado{qd fl_vdlg~: cjgl_ g wgogl_ jgpq_ 10$2013100000 kk, nmjd ")
        s = s & LD("0$20131000 #kk g kdl{wd nmjmaglz jgpq_, f_fmo 0$20131000 kk, l_gkdl{w")
        s = s & LD("#_~ cjgl_ ld `mj{wd cjglz jgpq_.")
    Case Else
        LangRu_6 = False
    End Select
End Function

Private Function LangEn(ByVal key As String) As String
    Dim s As String
    If LangEn_1(key, s) Then
    ElseIf LangEn_2(key, s) Then
    ElseIf LangEn_3(key, s) Then
    ElseIf LangEn_4(key, s) Then
    ElseIf LangEn_5(key, s) Then
    ElseIf LangEn_6(key, s) Then
    Else
        s = key
    End If
    LangEn = s
End Function

Private Function LangEn_1(ByVal key As String, s As String) As Boolean
    LangEn_1 = True
    s = ""
    Select Case key
    Case "app_title"
        s = s & LD("NeonCenterline")
    Case "center_title"
        s = s & LD("Centerline")
    Case "mode_mono"
        s = s & LD("Mono")
    Case "mode_script"
        s = s & LD("Script")
    Case "script_beta"
        s = s & LD("$000A$000A$26A0 Script mode is a work in progress and not stable yet: ki")
        s = s & LD("nks and stray pieces may appear at letter joints. Check the ")
        s = s & LD("lines before cutting.")
    Case "err_nodoc"
        s = s & LD("No document is open.")
    Case "err_nosel"
        s = s & LD("Select the letter outlines.")
    Case "err_noclosed"
        s = s & LD("The selection contains no closed outlines.")
    Case "err_fail"
        s = s & LD("Error: {0}")
    Case "err_nomodule"
        s = s & LD("Module $0022{0}$0022 was not found in the plugin project. Please rei")
        s = s & LD("nstall the plugin.")
    Case "done"
        s = s & LD("Done ({0}, cutter {1} mm): {2} objects, {3} lines created.$000AT")
        s = s & LD("ime {4} s (computing {5} s).")
    Case "warn_thin"
        s = s & LD("The cutter is too large for these letters: the stroke is $2248 {")
        s = s & LD("0} mm, but line ends stay {1} mm from the edge $2014 lines are s")
        s = s & LD("hortened, dropped entirely: {2}.$000AFor small text and engravin")
        s = s & LD("g, in the window before the run set the cutter diameter to t")
        s = s & LD("he engraver diameter, or $0022Line end inset from the edge$0022 to 0")
        s = s & LD("$20130.3 mm.")
    Case "prompt_tool"
        s = s & LD("Cutter / diffuser diameter, mm:$000A(a free line end is placed D")
        s = s & LD("/2 from the outline edge)")
    Case "bad_tool"
        s = s & LD("Invalid diameter.")
    Case "settings_title"
        s = s & LD("Settings")
    Case "lbl_tool"
        s = s & LD("Cutter diameter, mm:")
    Case "lbl_tool_hint"
        s = s & LD("A free line end is placed D/2 from the outline edge.")
    Case "chk_ask"
        s = s & LD("Ask for the parameters on every run")
    Case "lbl_lang"
        s = s & LD("Interface language:")
    Case "lang_auto"
        s = s & LD("Auto (as Windows)")
    Case "lang_ru"
        s = s & LD("#Orppigh")
    Case "lang_en"
        s = s & LD("English")
    Case "chk_upd"
        s = s & LD("Check for updates once a day")
    Case "lbl_version"
        s = s & LD("Version {0}")
    Case "btn_ok"
        s = s & LD("OK")
    Case "btn_cancel"
        s = s & LD("Cancel")
    Case "btn_update"
        s = s & LD("Check for updates")
    Case "btn_about"
        s = s & LD("About")
    Case "lang_changed"
        s = s & LD("The language applies to messages right away and to this wind")
        s = s & LD("ow next time it opens.")
    Case "upd_none"
        s = s & LD("You have the latest version ({0}).")
    Case "upd_found"
        s = s & LD("A new version {0} is available (you have {1}).$000A$000AWhat's new:$000A")
        s = s & LD("{2}$000A$000ADownload and install? The new version will start workin")
        s = s & LD("g after CorelDRAW is restarted.")
    Case "upd_ready"
        s = s & LD("Update {0} has been downloaded. It will be installed automat")
        s = s & LD("ically when you close CorelDRAW.")
    Case "upd_pending"
        s = s & LD("Update {0} is already downloaded and will be installed when ")
        s = s & LD("CorelDRAW is closed.")
    Case "upd_err"
        s = s & LD("Could not check for updates: {0}")
    Case "upd_bad"
        s = s & LD("The downloaded update did not pass the check: {0}. The updat")
        s = s & LD("e was cancelled $2014 you can install the new version from the a")
        s = s & LD("rchive on GitHub (install.cmd).")
    Case "upd_why_type"
        s = s & LD("it is not a plugin file (the network replaced or cut the dow")
        s = s & LD("nload)")
    Case "upd_why_sum"
        s = s & LD("checksum mismatch (expected {0}$2026, got {1}$2026)")
    Case "upd_why_nohash"
        s = s & LD("could not compute the checksum (certutil and PowerShell are ")
        s = s & LD("unavailable)")
    Case "upd_nourl"
        s = s & LD("The update address is not configured.")
    Case "upd_nogms"
        s = s & LD("The plugin file (.gms) to replace was not found.")
    Case Else
        LangEn_1 = False
    End Select
End Function

Private Function LangEn_2(ByVal key As String, s As String) As Boolean
    LangEn_2 = True
    s = ""
    Select Case key
    Case "about"
        s = s & LD("NeonCenterline {0}$000A$000ACenterline from letter outlines for mill")
        s = s & LD("ing flex-neon channels. Buttons: $0022Mono$0022 for monoline (neon) ")
        s = s & LD("fonts, $0022Script$0022 for handwritten fonts.$000A$000A$00A9 2026 SAB:CORE. Fre")
        s = s & LD("e for any use, including commercial. Selling or repackaging ")
        s = s & LD("the plugin is prohibited (see LICENSE).")
    Case "about_donate"
        s = s & LD("$000A$000AThe plugin is free. If it helps you, you can support the a")
        s = s & LD("uthor. Open the page?")
    Case "addon_done"
        s = s & LD("The toolbar is copied to$000A{0}$000A$000AFully restart CorelDRAW: the $0022")
        s = s & LD("Neon$0022 toolbar rebuilds itself at start, when a document open")
        s = s & LD("s or at the first plugin command. If the toolbar is not visi")
        s = s & LD("ble: Window > Toolbars > $0022#Ldml (#NeonCenterline)$0022.")
    Case "addon_fail"
        s = s & LD("The toolbar files were not copied to$000A{0}$000A$000A{1}$000A$000ARun the macro")
        s = s & LD(" NeonCenterline.NeonAddon.InstallToolbar again and confirm t")
        s = s & LD("he administrator prompt, or close CorelDRAW and copy the add")
        s = s & LD("on\NeonCenterline folder from the plugin archive there by ha")
        s = s & LD("nd.")
    Case "addon_nolog"
        s = s & LD("The copy did not start: the administrator prompt may have be")
        s = s & LD("en declined.")
    Case "err_nosel_any"
        s = s & LD("Nothing is selected.")
    Case "len_title"
        s = s & LD("Length of outlines and lines")
    Case "len_contours"
        s = s & LD("Outlines (cutting): {0} objects, {1} paths$000ATotal length: {2}")
    Case "len_neon"
        s = s & LD("Centerlines (neon): {0} pieces$000ATotal length: {1}$000ALongest pie")
        s = s & LD("ce: {2}")
    Case "len_none"
        s = s & LD("The selection has no curves to measure.")
    Case "rep_undo"
        s = s & LD("OK keeps the result. Cancel undoes all of it (like Ctrl+Z).")
    Case "rep_undo_fail"
        s = s & LD("Could not undo $2014 press Ctrl+Z.")
    Case "rep_kept"
        s = s & LD("Originals kept: the result is copies on top of them.")
    Case "chk_keep"
        s = s & LD("Keep the original (the result is a copy on top)")
    Case "lbl_keep_orig"
        s = s & LD("Keep the original:")
    Case "chk_keep_clean"
        s = s & LD("clean")
    Case "chk_keep_split"
        s = s & LD("split")
    Case "chk_keep_simp"
        s = s & LD("simplify")
    Case "prog_start"
        s = s & LD("Preparing$2026")
    Case "prog_step"
        s = s & LD("{0}: {1} of {2}")
    Case "prog_time"
        s = s & LD("Working: {0} s")
    Case "prog_read"
        s = s & LD("Reading outlines")
    Case "prog_core"
        s = s & LD("Centerlines, letter groups")
    Case "prog_emit"
        s = s & LD("Drawing the lines")
    Case "prog_objects"
        s = s & LD("Objects")
    Case "prog_compare"
        s = s & LD("Comparing shapes")
    Case "prog_nest_geo"
        s = s & LD("Part outlines ({0})")
    Case "prog_nest_pack"
        s = s & LD("Nesting, variant {0} of {1}")
    Case "prog_nest_place"
        s = s & LD("Parts onto sheets ({0})")
    Case "prog_count"
        s = s & LD("{0}: {1}")
    Case "prog_measure"
        s = s & LD("Measuring")
    Case "prog_find"
        s = s & LD("Looking for centerlines")
    Case "prog_cuts"
        s = s & LD("Cuts")
    Case "prog_collect"
        s = s & LD("Collecting objects")
    Case "prog_groups"
        s = s & LD("Groups")
    Case "prog_cutfit"
        s = s & LD("Fitting to the cut step")
    Case "sel_none"
        s = s & LD("There are no plugin centerlines on this page.")
    Case "unit_mm"
        s = s & LD("mm")
    Case "unit_m"
        s = s & LD("m")
    Case "cut_title"
        s = s & LD("Cut step")
    Case Else
        LangEn_2 = False
    End Select
End Function

Private Function LangEn_3(ByVal key As String, s As String) As Boolean
    LangEn_3 = True
    s = ""
    Select Case key
    Case "cut_nostep"
        s = s & LD("The cut step is not set (Settings).")
    Case "cut_report"
        s = s & LD("Fit to cut step {0} mm: {1} pieces shortened ({2} mm removed")
        s = s & LD(" in total), {3} already fit, {4} not fitted (closed, no free")
        s = s & LD(" ends, or shorter than one step).")
    Case "cut_sum_classic"
        s = s & LD("Neon in cut steps of {0} mm: {1} mm, remainders (not filled ")
        s = s & LD("by neon) {2} mm")
    Case "cut_sum_neon2"
        s = s & LD("LED strip in cut steps of {0} mm: {1} mm, groove without str")
        s = s & LD("ip {2} mm")
    Case "lbl_type"
        s = s & LD("Neon type:")
    Case "type_classic"
        s = s & LD("LED neon (strip inside)")
    Case "type_neon2"
        s = s & LD("Neon 2.0 (diffuser in a groove, separate strip)")
    Case "lbl_cut"
        s = s & LD("Cut step, mm:")
    Case "chk_fit"
        s = s & LD("Fit lines to the cut step (shorten free ends)")
    Case "lbl_trim"
        s = s & LD("Shorten:")
    Case "trim_both"
        s = s & LD("equally at both ends")
    Case "trim_end"
        s = s & LD("only at the line end")
    Case "chk_ask_run"
        s = s & LD("Show this window on every run (turn back on in Settings)")
    Case "btn_run"
        s = s & LD("Build")
    Case "upd_net_hint"
        s = s & LD("$000A$000AGitHub may be unreachable from your network. You can downl")
        s = s & LD("oad the new version manually: github.com/ShiZa039/NeonCenter")
        s = s & LD("line/releases")
    Case "lbl_ref"
        s = s & LD("Script $2014 font as a reference (cutter ignored):")
    Case "ref_off"
        s = s & LD("no: mill the outline itself")
    Case "ref_tip"
        s = s & LD("yes, line to the stroke tip")
    Case "ref_half"
        s = s & LD("yes, inset by half the stroke width")
    Case "lbl_endoff"
        s = s & LD("Line end inset from the edge, mm:")
    Case "endoff_default"
        s = s & LD("D/2 (cutter radius)")
    Case "fallback_lang"
        s = s & LD("Interface language: Yes $2014 Russian, No $2014 English, Cancel $2014 au")
        s = s & LD("to.")
    Case "fallback_upd"
        s = s & LD("Check for updates once a day?")
    Case "fallback_ask"
        s = s & LD("Ask for the diameter on every run?")
    Case "clean_title"
        s = s & LD("Clean outline")
    Case "clean_none"
        s = s & LD("The selection has no curves or text (rectangles and ellipses")
        s = s & LD(" are skipped $2014 they have no overlaps).")
    Case "clean_nochange"
        s = s & LD("No overlaps or loops found $2014 nothing was changed.")
    Case "clean_report"
        s = s & LD("Cleaned {0} of {1} objects.$000AOutline crossings removed: {2}.$000A")
        s = s & LD("Closed outlines: {3} before, {4} after.$000A$000AEach object is one ")
        s = s & LD("curve.")
    Case "clean_skipped"
        s = s & LD("Skipped {0} objects that are neither curves nor text.")
    Case "simp_title"
        s = s & LD("Simplify outline")
    Case "simp_none"
        s = s & LD("The selection has no curves. Convert text and shapes to curv")
        s = s & LD("es first (Ctrl+Q).")
    Case "simp_report"
        s = s & LD("Tolerance {0} mm: {1} of {2} curves simplified.$000ANodes: {3} b")
        s = s & LD("efore, {4} after ({5}% fewer).")
    Case "simp_skipped"
        s = s & LD("Skipped {0} objects that are not curves (convert to curves: ")
        s = s & LD("Ctrl+Q).")
    Case "lbl_simp"
        s = s & LD("Simplify outline, tolerance, mm:")
    Case "prompt_simp"
        s = s & LD("Simplify outline: tolerance, mm$000A(how far the curve may devia")
        s = s & LD("te from the original; 0.1 is invisible)")
    Case "bad_simp"
        s = s & LD("The simplify tolerance must be a number from 0.01 to 5 mm.")
    Case "lbl_simp_tol"
        s = s & LD("Tolerance, mm:")
    Case "simp_hint"
        s = s & LD("How far the curve may deviate from the original. 0.1 mm is i")
        s = s & LD("nvisible to the eye and the cutter; more smooths ragged trac")
        s = s & LD("ed outlines harder.")
    Case "btn_simp"
        s = s & LD("Simplify")
    Case "chk_simp_ask"
        s = s & LD("Show the simplify window on every run")
    Case Else
        LangEn_3 = False
    End Select
End Function

Private Function LangEn_4(ByVal key As String, s As String) As Boolean
    LangEn_4 = True
    s = ""
    Select Case key
    Case "tb_outdated"
        s = s & LD("The $0022Neon$0022 toolbar in CorelDRAW is outdated: this plugin ver")
        s = s & LD("sion has new buttons on it. Update the toolbar now?$000A$000AWindows")
        s = s & LD(" will ask for administrator rights. Then fully restart Corel")
        s = s & LD("DRAW and the new buttons appear by themselves.")
    Case "tb_reset_done"
        s = s & LD("The $0022Neon$0022 toolbar is reset to the button set of this plugin")
        s = s & LD(" version.$000A$000A{0}")
    Case "tb_reset_none"
        s = s & LD("Could not reset the $0022Neon$0022 toolbar.$000A$000A{0}$000AIf the toolbar is n")
        s = s & LD("ot listed (Window > Toolbars), install it: macro NeonCenterl")
        s = s & LD("ine.NeonAddon.InstallToolbar, then fully restart CorelDRAW.")
    Case "tb_files_old"
        s = s & LD("The $0022Neon$0022 toolbar files in the CorelDRAW Addons folder are ")
        s = s & LD("from an older plugin version (or missing):$000A{0}$000AButton set [N")
        s = s & LD("eonToolbar {1}] is needed.$000A$000AUpdate the toolbar files now? Wi")
        s = s & LD("ndows will ask for administrator rights; then fully restart ")
        s = s & LD("CorelDRAW.")
    Case "grom_title"
        s = s & LD("Grommets")
    Case "grom_nosel"
        s = s & LD("Select the banner (backing): a rectangle or a closed outline")
        s = s & LD(".")
    Case "grom_none"
        s = s & LD("The selection has no closed outline that fits grommets at th")
        s = s & LD("is edge distance.")
    Case "grom_report"
        s = s & LD("Grommets: {0} (objects: {1}).$000AEdge to grommet centre {2} mm,")
        s = s & LD(" distance between grommets at most {3} mm, point $00D8{4} mm.$000APo")
        s = s & LD("ints are on the $0022Grommets$0022 layer, one group per object.")
    Case "grom_report_reinf"
        s = s & LD("Corner reinforcement: one more point {0} mm from each corner")
        s = s & LD(".")
    Case "grom_skipped"
        s = s & LD("Objects skipped: {0} (no closed outline, or the edge distanc")
        s = s & LD("e is larger than the object).")
    Case "grom_why_nocurve"
        s = s & LD("could not get the outline")
    Case "grom_why_open"
        s = s & LD("no closed outline ({0} subpaths, {1} closed; the ends do not")
        s = s & LD(" meet)")
    Case "grom_why_small"
        s = s & LD("grommets do not fit at this edge distance")
    Case "grom_t_curve"
        s = s & LD("curve")
    Case "grom_t_rect"
        s = s & LD("rectangle")
    Case "grom_t_ellipse"
        s = s & LD("ellipse")
    Case "grom_t_poly"
        s = s & LD("polygon")
    Case "grom_t_text"
        s = s & LD("text")
    Case "grom_t_other"
        s = s & LD("object (type {0})")
    Case "grom_inner"
        s = s & LD("Inside the banner (artwork, text, frames) $2014 no grommets: {0}")
        s = s & LD(".")
    Case "grom_layer"
        s = s & LD("Grommets")
    Case "grom_dot"
        s = s & LD("Grommet")
    Case "grom_group"
        s = s & LD("Grommets: {0}")
    Case "lbl_grom_edge"
        s = s & LD("Edge to grommet centre, mm:")
    Case "lbl_grom_step"
        s = s & LD("Distance between grommets, max, mm:")
    Case "lbl_grom_dot"
        s = s & LD("Point diameter, mm:")
    Case "lbl_grom_color"
        s = s & LD("Point colour:")
    Case "grom_c_black"
        s = s & LD("Black (K100)")
    Case "grom_c_red"
        s = s & LD("Red (M100 Y100)")
    Case "grom_c_magenta"
        s = s & LD("Magenta (M100)")
    Case "grom_c_blue"
        s = s & LD("Blue (C100 M100)")
    Case "grom_c_green"
        s = s & LD("Green (C100 Y100)")
    Case "grom_c_white"
        s = s & LD("White")
    Case "grom_c_other"
        s = s & LD("Other$2026")
    Case "chk_grom_reinf"
        s = s & LD("Corner reinforcement: extra point at, mm")
    Case "chk_grom_hem"
        s = s & LD("Hem outline, allowance, mm:")
    Case "chk_grom_layer"
        s = s & LD("Grommets on a separate $0022Grommets$0022 layer (otherwise on the ba")
        s = s & LD("nner layer)")
    Case "bad_grom_hem"
        s = s & LD("The hem allowance must be between 0 and 1000 mm.")
    Case "grom_hem_name"
        s = s & LD("Hem {0} mm")
    Case "grom_hem_rep"
        s = s & LD("Hem outline: {0} mm outside the edge ({1} objects) $2014 on the ")
        s = s & LD("banner layer.")
    Case Else
        LangEn_4 = False
    End Select
End Function

Private Function LangEn_5(ByVal key As String, s As String) As Boolean
    LangEn_5 = True
    s = ""
    Select Case key
    Case "grom_hint"
        s = s & LD("A grommet in every corner, evenly along the sides, never fur")
        s = s & LD("ther apart than the given distance. A rounded corner counts ")
        s = s & LD("as a corner; circles and ovals get grommets evenly all round")
        s = s & LD(". The hem outline is a cut line the allowance outside the ba")
        s = s & LD("nner edge (grommets are still measured from the edge).")
    Case "btn_grom"
        s = s & LD("Place")
    Case "bad_grom"
        s = s & LD("Check the values: edge 0$2013500 mm, distance 20$201310000 mm, point")
        s = s & LD(" 0$2013100 mm.")
    Case "split_title"
        s = s & LD("Split into letters")
    Case "split_none"
        s = s & LD("The selection has no curves or text made of several letters ")
        s = s & LD("$2014 nothing to split.")
    Case "split_report"
        s = s & LD("Objects split: {0}, separate letters and signs: {1}. Holes s")
        s = s & LD("tay in their letters, dots and accents ($0401, #H, #i, !, ?) stay ")
        s = s & LD("with their letter; parts inside holes (the letter in $00AE, the ")
        s = s & LD("dot in a zero) are separate objects. Single-letter objects l")
        s = s & LD("eft as they were: {2}.$000A$000AIf letters overlap, run $0022Clean outli")
        s = s & LD("ne$0022 first.")
    Case "noodle_title"
        s = s & LD("Strips")
    Case "lbl_noodle_len"
        s = s & LD("Strip length, mm:")
    Case "lbl_noodle_wid"
        s = s & LD("Strip width, mm:")
    Case "lbl_noodle_cnt"
        s = s & LD("Number of strips:")
    Case "lbl_noodle_overh"
        s = s & LD("Tails of horizontal cuts, mm:")
    Case "lbl_noodle_overv"
        s = s & LD("Tails of vertical end cuts, mm:")
    Case "chk_noodle_tool"
        s = s & LD("Router $2014 cutter diameter, mm:")
    Case "bad_noodle_tool"
        s = s & LD("The cutter diameter must be between 0 and 100 mm.")
    Case "noodle_hint"
        s = s & LD("Strips for a plotter, laser or router. The cutting order is ")
        s = s & LD("the object order: horizontal cuts top to bottom in a zigzag ")
        s = s & LD("(directions alternate, no idle runs), and at the very end of")
        s = s & LD(" the list two vertical end cuts, also along the way. With $0022R")
        s = s & LD("outer$0022 the cuts are cutter centre lines: horizontal ones are")
        s = s & LD(" spaced by width + diameter, end cuts sit a radius past the ")
        s = s & LD("strip ends, so after the router the strips come out exactly ")
        s = s & LD("the size you set. Tails run a cut past the other cut's line ")
        s = s & LD("so each strip comes off completely: 1$20132 mm for a laser or pl")
        s = s & LD("otter, 2$20135 mm of margin for a router. One group in the middl")
        s = s & LD("e of the page.")
    Case "btn_noodle"
        s = s & LD("Make strips")
    Case "bad_noodle"
        s = s & LD("Check the values: length 0$2013100000 mm, width 0$201310000 mm, coun")
        s = s & LD("t a whole number 1$20131000, tails 0$20131000 mm.")
    Case "noodle_group"
        s = s & LD("Strips: {0} pcs $00B7 {1} $00D7 {2} mm")
    Case "noodle_report"
        s = s & LD("Strips: {0} strips {1} $00D7 {2} mm.$000ACuts: {3} horizontal ones t")
        s = s & LD("op to bottom in a zigzag, then 2 vertical end cuts at the en")
        s = s & LD("d of the list; tails: horizontal {4} mm, vertical {5} mm. Si")
        s = s & LD("ze with tails: {6} $00D7 {7} mm.")
    Case "noodle_rep_tool"
        s = s & LD("Router $2300{0} mm: the cuts are cutter centre lines, horizontal")
        s = s & LD(" ones {1} mm apart, end cuts {2} mm apart $2014 strips after the")
        s = s & LD(" router are {3} $00D7 {4} mm.")
    Case "noodle_rep_tail"
        s = s & LD("One group in the middle of the page; the object order is the")
        s = s & LD(" cutting order.")
    Case "same_title"
        s = s & LD("Group identical")
    Case "same_nosel"
        s = s & LD("Select the objects: curves, text, shapes.")
    Case "same_none"
        s = s & LD("No identical objects found: all {0} are different.$000A$000AIdentica")
        s = s & LD("l means the same shape within 0.1 mm (no rotation or mirrori")
        s = s & LD("ng) and the same colour.")
    Case "same_report"
        s = s & LD("Objects: {0}, different shapes: {1}.$000AIdentical ones are grou")
        s = s & LD("ped: {2} groups with {3} objects. Without a pair: {4}.$000A$000AIden")
        s = s & LD("tical means the same shape within 0.1 mm (no rotation or mir")
        s = s & LD("roring) and the same colour.")
    Case "same_list"
        s = s & LD("Shapes (count $2014 size $2014 colour):")
    Case "same_line"
        s = s & LD("{0} pcs $2014 {1} $00D7 {2} mm $2014 {3}")
    Case "same_more"
        s = s & LD("$2026 and {0} more shapes")
    Case "same_group"
        s = s & LD("Identical: {0} pcs $00B7 {1} $00D7 {2} mm")
    Case "same_failed"
        s = s & LD("Could not group {0} sets (objects on different layers?).")
    Case "no_color"
        s = s & LD("no colour")
    Case "nest_title"
        s = s & LD("Nest on sheets (beta)")
    Case "nest_beta"
        s = s & LD("$26A0 Nesting is a work in progress: by outline on a grid of abo")
        s = s & LD("ut 1 mm (coarser on large sheets); parts may go into letter ")
        s = s & LD("holes and inside other parts (the $0022nested$0022 box in the window")
        s = s & LD("). A part's outline must be closed, otherwise another part m")
        s = s & LD("ay go inside it even without nesting. Check the layout befor")
        s = s & LD("e cutting.")
    Case "nest_nosel"
        s = s & LD("Select the parts to nest.")
    Case "nest_layer"
        s = s & LD("Nesting")
    Case "nest_frame"
        s = s & LD("Nesting: sheet frame {0}")
    Case "nest_labname"
        s = s & LD("Nesting: sheet label {0}")
    Case "nest_group"
        s = s & LD("Nesting: sheet {0} of {1} $00B7 {2}")
    Case "nest_label"
        s = s & LD("Sheet {0} of {1} $00B7 {2} $00D7 {3} mm $00B7 {4} $00B7 {5} parts $00B7 {6}% fil")
        s = s & LD("led")
    Case "nest_report"
        s = s & LD("Parts nested: {0}, sheets: {1}.$000A{2}Edge margin {3} mm, gap b")
        s = s & LD("etween parts {4} mm, {5}.")
    Case Else
        LangEn_5 = False
    End Select
End Function

Private Function LangEn_6(ByVal key As String, s As String) As Boolean
    LangEn_6 = True
    s = ""
    Select Case key
    Case "nest_how_c0"
        s = s & LD("by outline, no rotation")
    Case "nest_how_c1"
        s = s & LD("by outline, rotation by 180$00B0")
    Case "nest_how_c2"
        s = s & LD("by outline, rotation in 90$00B0 steps")
    Case "nest_how_box"
        s = s & LD("by bounding boxes, no rotation")
    Case "nest_how_noin"
        s = s & LD(", not inside other parts")
    Case "nest_rep_sheet"
        s = s & LD("Sheet {0} $00D7 {1} mm (length $00D7 width). ")
    Case "nest_rep_auto"
        s = s & LD("Sheet width {0} mm, length to fit the parts ({1}$2013{2} mm, rou")
        s = s & LD("nded up to 10 mm). ")
    Case "nest_rep_color"
        s = s & LD("{0}: {1} sheets, {2} parts, {3}% filled")
    Case "nest_rep_widths"
        s = s & LD(" (length: {0})")
    Case "nest_rep_tail"
        s = s & LD("Each sheet is a group to the right of the selection on the p")
        s = s & LD("arts' layer: frame, label and parts (labels can be turned of")
        s = s & LD("f in the window). Fill is by part bounding boxes.")
    Case "nest_rep_tail_c"
        s = s & LD("Each sheet is a group to the right of the selection on the p")
        s = s & LD("arts' layer: frame, label and parts (labels can be turned of")
        s = s & LD("f in the window). Fill is by part area (holes excluded).")
    Case "nest_unfit"
        s = s & LD("Larger than the sheet working area, left in place: {0}.")
    Case "nest_none"
        s = s & LD("No part fits on the sheet with these margins.")
    Case "all_colors"
        s = s & LD("all colours")
    Case "lbl_nest_size"
        s = s & LD("Preset size:")
    Case "lbl_nest_len"
        s = s & LD("Sheet length, mm:")
    Case "lbl_nest_wid"
        s = s & LD("Sheet width, mm:")
    Case "prompt_nest_size"
        s = s & LD("Sheet: length $00D7 width, mm (e.g. 900x600):")
    Case "lbl_nest_margin"
        s = s & LD("Margin from the sheet edge, mm:")
    Case "lbl_nest_gap"
        s = s & LD("Gap between parts, mm:")
    Case "chk_nest_color"
        s = s & LD("Each colour on its own sheets")
    Case "lbl_nest_method"
        s = s & LD("Nesting:")
    Case "nest_m_contour"
        s = s & LD("by outline $2014 tighter")
    Case "nest_m_rot180"
        s = s & LD("by outline, rotate by 180$00B0")
    Case "nest_m_rot90"
        s = s & LD("by outline, rotate in 90$00B0 steps")
    Case "nest_m_box"
        s = s & LD("by bounding boxes $2014 faster")
    Case "chk_nest_groups"
        s = s & LD("Groups stay whole (each group is one part)")
    Case "chk_nest_keep"
        s = s & LD("Keep the original $2014 copies of the parts go onto the sheets")
    Case "nest_rep_kept"
        s = s & LD("The originals stay in place; the sheets hold copies.")
    Case "chk_nest_inside"
        s = s & LD("Parts inside other parts (nested): letter holes, frames")
    Case "lbl_nest_fill"
        s = s & LD("Fill the sheet:")
    Case "nest_fill_cols"
        s = s & LD("in columns $2014 top to bottom")
    Case "nest_fill_rows"
        s = s & LD("in rows $2014 left to right")
    Case "nest_fill_auto"
        s = s & LD("as tight as possible (auto)")
    Case "chk_nest_labels"
        s = s & LD("Sheet labels (above the frame, in the sheet group)")
    Case "chk_nest_auto"
        s = s & LD("Length to fit the parts, at least, mm:")
    Case "nest_hint"
        s = s & LD("Beta. Length is horizontal, width is vertical; any size: typ")
        s = s & LD("e your own (laser bed, offcut) and it appears in the presets")
        s = s & LD(". By outline is tighter: parts fit into each other, and with")
        s = s & LD(" $0022nested$0022 also into letter holes and frames; rotate only if ")
        s = s & LD("the material allows it (no texture or grain). By bounding bo")
        s = s & LD("xes is faster. No part is closer than the margin to the edge")
        s = s & LD(" or the gap to another part; for a router, use a gap of at l")
        s = s & LD("east the cutter diameter. With $0022Length to fit$0022 each sheet is")
        s = s & LD(" trimmed to the parts. Groups stay whole (plates, assemblies")
        s = s & LD("); untick it to split groups into their elements $2014 handy for")
        s = s & LD(" an SVG/DXF import that comes as one group. $0022Identical$0022 grou")
        s = s & LD("ps are always nested one by one.")
    Case "btn_nest"
        s = s & LD("Nest")
    Case "bad_nest"
        s = s & LD("Check the values: sheet length and width 10$2013100000 mm, margi")
        s = s & LD("n 0$20131000 mm and under half the sheet, gap 0$20131000 mm, minimum")
        s = s & LD(" length not above the sheet length.")
    Case Else
        LangEn_6 = False
    End Select
End Function
