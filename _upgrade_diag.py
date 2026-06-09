#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""智能诊断模块升级脚本 V2"""

import os

# 读取原文件
with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 新的Tab 6代码（增强版）
new_tab6 = r'''# ═══════════════════════════════════════════════════════════════
# TAB 6: 智能诊断 V2（多因子归因 + 健康评分 + 可执行措施 + 正向亮点）
# ═══════════════════════════════════════════════════════════════
with tabs[5]:
    st.markdown('<div class="section-title">🔍 智能问题定位诊断 & 优化措施</div>', unsafe_allow_html=True)
    st.caption(f'诊断区间：{s} ~ {e} | 对比区间：相同天数的上一期 | 下钻维度：渠道 → 品类 → 型号')

    # ══════════════════════════════════════
    # A. 核心数据准备
    # ══════════════════════════════════════
    cur_days = (end - start).days + 1
    mom_end_dt = start - datetime.timedelta(days=1)
    mom_start_dt = mom_end_dt - datetime.timedelta(days=cur_days - 1)
    prev_s_d = str(mom_start_dt)
    prev_e_d = str(mom_end_dt)

    cur_sum = summarize(daily)

    def _agg_by_dims(rows, dims):
        out = {}
        for r in rows:
            key = tuple(r.get(d) or '未标注' for d in dims)
            out.setdefault(key, {m: 0.0 for m in METRICS})
            for m in METRICS:
                out[key][m] += float(r.get(m, 0) or 0)
        for k, v in out.items():
            v['支付转化率'] = v['支付买家数'] / v['商品访客数'] if v['商品访客数'] else 0
            v['客单价'] = v['支付金额'] / v['支付买家数'] if v['支付买家数'] else 0
            v['退款率'] = v['成功退款金额'] / v['支付金额'] if v['支付金额'] else 0
        return out

    cur_rows_all = [r for r in data['daily']]
    prev_rows_all = []
    for r in data['daily']:
        d = r.get('日期', '')
        if len(d) == 7: d = d + '-01'
        if prev_s_d <= d <= prev_e_d: prev_rows_all.append(r)

    cur_by_model   = _agg_by_dims(cur_rows_all, ['渠道','品类','型号'])
    cur_by_cat     = _agg_by_dims(cur_rows_all, ['渠道','品类'])
    cur_by_channel = _agg_by_dims(cur_rows_all, ['渠道'])
    prev_by_model  = _agg_by_dims(prev_rows_all, ['渠道','品类','型号'])
    prev_by_cat    = _agg_by_dims(prev_rows_all, ['渠道','品类'])
    prev_by_channel= _agg_by_dims(prev_rows_all, ['渠道'])

    prev_sum_all = summarize(prev_rows_all)
    def _global_chg(k):
        c = cur_sum.get(k, 0); p = prev_sum_all.get(k, 0)
        return (c - p) / p if p else None

    gmv_g = _global_chg('支付金额')
    vis_g = _global_chg('商品访客数')
    cvr_g = _global_chg('支付转化率')
    aov_g = _global_chg('客单价')
    ref_g = _global_chg('退款率')

    WARN_T = -0.05; DANGER_T = -0.15

    def _pct(v): return f'{v*100:+.1f}%' if v is not None else '--'

    # ──────────────────────────────────────
    # A0. 总体健康评分 & 一句话结论（新增）
    # ──────────────────────────────────────
    scores = {}
    for name, val in [('GMV',gmv_g),('流量',vis_g),('转化率',cvr_g),('客单价',aov_g),('退款率',ref_g)]:
        if val is None: scores[name] = 100
        elif val > WARN_T: scores[name] = 100
        elif val > DANGER_T: scores[name] = 60 + int((val - WARN_T) / (0 - WARN_T) * 40)
        else: scores[name] = max(0, int(val / DANGER_T * 60))
    if ref_g is not None and ref_g > 0:
        scores['退款率'] = max(0, 100 - abs(ref_g) * 300)
    health_score = sum(scores.values()) / len(scores)

    if health_score >= 90: hv = ('🟢 整体健康', '#22c55e', '各项核心指标表现良好，继续保持现有经营策略。')
    elif health_score >= 70: hv = ('🟡 需要关注', '#f59e0b', f'部分指标出现波动（综合得分{health_score:.0f}/100），建议重点关注下方标红项。')
    elif health_score >= 50: hv = ('🠤 存在风险', '#ef4444', f'多项指标明显下滑（综合得分{health_score:.0f}/100），建议立即执行P0优先级行动。')
    else: hv = ('⚠️ 紧急告警', '#dc2626', f'整体经营状况堪忧（综合得分{health_score:.0f}/100），请优先处理所有P0任务！')

    sc1, sc2, sc3 = st.columns([1, 4, 2])
    with sc1:
        st.metric('健康评分', f'{health_score:.0f}', delta=None,
                  help=f"GMV:{scores['GMV']} | 流量:{scores['流量']} | 转化:{scores['转化率']} | 客单价:{scores['客单价']} | 退款:{scores['退款率']}")
    with sc2:
        st.markdown(f"<div style='background:{hv[1]}15;border-left:4px solid {hv[1]};border-radius:8px;padding:12px 16px;margin-top:16px;'>"
                     f"<strong>{hv[0]}</strong>&nbsp;&nbsp;{hv[2]}</div>", unsafe_allow_html=True)
    with sc3:
        st.metric('异常型号数', f'{len([m for m in cur_by_model.values()])}' if False else '-', delta=None)

    # ══════════════════════════════════════
    # B. 第一层：全局健康度总览（5卡片 + 归因提示）
    # ══════════════════════════════════════
    s1,s2,s3,s4,s5 = st.columns(5)
    metrics_info = [
        ('\U0001F4B0 支付金额',gmv_g,cur_sum.get('支付金额',0),prev_sum_all.get('支付金额',0),'¥',False),
        ('\U0001F441 商品访客数',vis_g,cur_sum.get('商品访客数',0),prev_sum_all.get('商品访客数',0),'',False),
        ('\U0001F504 支付转化率',cvr_g,cur_sum.get('支付转化率',0)*100,prev_sum_all.get('支付转化率',0)*100,'',True),
        ('\U0001F3AB 客单价',aov_g,cur_sum.get('客单价',0),prev_sum_all.get('客单价',0),'¥',False),
       ('\U00021A9\U000FE0F 退款率',ref_g,cur_sum.get('退款率',0)*100,prev_sum_all.get('退款率',0)*100,'',True),
    ]
    for col,(mname,mch,cv,pv,pre,ispct) in zip([s1,s2,s3,s4,s5],metrics_info):
        with col:
            lvl='ok' if mch is None or mch>WARN_T else ('warn' if mch>DANGER_T else 'danger')
            icon={'danger':'\U0001F534','warn':'\U0001F7E1','ok':'\U0001F7E2'}[lvl]
            cv_s=f'{cv:,.0f}' if not ispct else f'{cv:.2f}%'
            pv_s=f'{pv:,.0f}' if not ispct else f'{pv:.2f}%'
            ch_s=_pct(mch)
            bg={'danger':'#fef2f2','warn':'#fff7ed','ok':'#f0fdf4'}[lvl]
            border={'danger':'#fca5a5','warn':'#fdba74','ok':'#86efac'}[lvl]
            cause_hint = ''
            if '支付金额' in mname and mch is not None and mch < 0:
                parts=[]
                if vis_g and vis_g<-0.03: parts.append(f'流量{_pct(vis_g)}')
                if cvr_g and cvr_g<-0.02: parts.append(f'转化{_pct(cvr_g)}')
                if aov_g and aov_g<-0.02: parts.append(f'客单价{_pct(aov_g)}')
                cause_hint=f'<div style="font-size:10px;color:#ea580c;margin-top:2px;">主因：{"+".join(parts) if parts else "多因子"}</div>'
            elif '访客' in mname and mch is not None and mch < -0.05:
                cause_hint='<div style="font-size:10px;color:#ea580c;margin-top:2px;">→ 可能原因：推广降权/搜索排名下降/季节性波动</div>'
            elif '转化' in mname and mch is not None and mch < -0.03:
                cause_hint='<div style="font-size:10px;color:#ea580c;margin-top:2px;">→ 可能原因：价格竞争力下降/差评累积/详情页体验差</div>'
            st.markdown(
                f'<div style="background:{bg};border:1px solid {border};border-radius:14px;padding:12px;text-align:center;">'
                f'<div style="font-size:11.5px;color:#64748b;font-weight:700;">{mname}</div>'
                f'<div style="font-size:21px;font-weight:900;color:#0f172a;margin:4px 0;">{pre}{cv_s}</div>'
                f'<div style="font-size:11px;color:#94a3b8;">vs 上期 {pre}{pv_s} ({ch_s})</div>{cause_hint}</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════
    # C. 第二层：根因下钻 — 型号级定位（增强归因）
    # ══════════════════════════════════════
    st.markdown('<hr style="margin:18px 0;border:none;border-top:1px dashed #cbd5e1;">')
    st.markdown('<div class="section-title">\U0001F3AF 根因定位：下钻到型号级</div>', unsafe_allow_html=True)
    st.caption('展示各维度中「环比下滑最严重」的具体型号。归因列使用多因子加权分析，自动识别主要衰退驱动因素。')

    # --- C1. 各渠道内 GMV 下滑最严重的 TOP 型号 ---
    st.markdown('#### \U0001F4CD 各渠道内 GMV 下滑最严重的 TOP 型号（含多因子归因）')
    ch_model_issues=[]
    total_cur_gmv = cur_sum.get('支付金额', 1)
    for ch_key,cur_v in cur_by_channel.items():
        ch_name=ch_key[0]
        prev_v=prev_by_channel.get(ch_key,{})
        ch_gmv_c=cur_v.get('支付金额',0); ch_gmv_p=prev_v.get('支付金额',0)
        ch_chg=(ch_gmv_c-ch_gmv_p)/ch_gmv_p if ch_gmv_p else None
        if ch_chg and ch_chg<0:
            worst_models=[]
            for mk_key,mv in cur_by_model.items():
                if mk_key[0]!=ch_name: continue
                pv_m=prev_by_model.get(mk_key,{})
                mc=mv.get('支付金额',0); mp=pv_m.get('支付金额',0)
                m_chg=(mc-mp)/mp if mp else None
                if m_chg is not None and m_chg<0:
                    worst_models.append({
                        '渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                        '本期GMV':mc,'上期GMV':mp,'环比':m_chg,
                        '本期转化率':mv.get('支付转化率',0),'上期转化率':pv_m.get('支付转化率',0),
                        '本期访客':mv.get('商品访客数',0),'上期访客':pv_m.get('商品访客数',0),
                        '本期件数':mv.get('支付件数',0),'上期件数':pv_m.get('支付件数',0),
                        '本期客单价':mv.get('客单价',0),'上期客单价':pv_m.get('客单价',0),
                    })
            worst_models.sort(key=lambda x:x['环比'])
            for wm in worst_models[:3]: ch_model_issues.append(wm)

    if ch_model_issues:
        tbl_rows=[]
        for idx,w in enumerate(ch_model_issues[:25]):
            vis_c=w['本期访客'];vis_p=w['上期访客']
            vis_chg_m=(vis_c-vis_p)/vis_p if vis_p else None
            cvr_c=w['本期转化率'];cvr_p=w['上期转化率']
            cvr_diff=(cvr_c-cvr_p)*100 if cvr_p else None
            aov_c=w['本期客单价'];aov_p=w['上期客单价']
            aov_chg_m=(aov_c-aov_p)/aov_p if aov_p else None
            factors=[]
            if vis_chg_m is not None:
                if vis_chg_m<=-0.20: factors.append(('流量崩塌',f'\u2193{_pct(vis_chg_m)}',3))
                elif vis_chg_m<=-0.10: factors.append(('流量大幅下降',f'\u2193{_pct(vis_chg_m)}',2))
                elif vis_chg_m<=-0.05: factors.append(('流量小幅下滑',f'\u2193{_pct(vis_chg_m)}',1))
            if cvr_diff is not None:
                if cvr_diff<=-5: factors.append(('转化率崩溃',f'\u2193{cvr_diff:.1f}pp',3))
                elif cvr_diff<=-2: factors.append(('转化率明显下降',f'\u2193{cvr_diff:.1f}pp',2))
                elif cvr_diff<=-0.5: factors.append(('转化率微降',f'\u2193{cvr_diff:.1f}pp',1))
            if aov_chg_m is not None:
                if aov_chg_m<=-0.20: factors.append(('客单价暴跌',f'\u2193{_pct(aov_chg_m)}',3))
                elif aov_chg_m<=-0.10: factors.append(('客单价明显下跌',f'\u2193{_pct(aov_chg_m)}',2))
                elif aov_chg_m<=-0.03: factors.append(('客单价微跌',f'\u2193{_pct(aov_chg_m)}',1))
            factors.sort(key=lambda x:x[2], reverse=True)
            if factors:
                reason_parts=[]
                for fname,fdetail,fw in factors[:3]:
                    c='#dc2626' if fw==3 else ('#ea580c' if fw==2 else '#f59e0b')
                    reason_parts.append(f"<span style='color:{c}'>\u25CF {fname}:{fdetail}</span>")
                reason_str=' '.join(reason_parts)
            else:
                reason_str='<span style=\'color:#64748b\'>多因素平稳下滑</span>'
            impact_amt=max(0, w['上期GMV']-w['本期GMV'])
            impact_pct=(impact_amt/total_cur_gmv*100) if total_cur_gmv else 0
            tbl_rows.append({
                '序号':idx+1,
                '严重度':'\U0001F534' if w['环比']<DANGER_T else ('\U0001F7E0' if w['环比']<WARN_T else '\U0001F7E1'),
                '渠道':w['渠道'],'品类':w['品类'],'型号':w['型号'],
                '本期GMV':f"\u00A5{w['本期GMV']:,.0f}",'上期GMV':f"\u00A5{w['上期GMV']:,.0f}",
                '环比变化':f"<span style='color:#dc2626;font-weight:700'>{_pct(w['环比'])}</span>",
                '拖累总额':f"\u00A5{-impact_amt:,.0f}({impact_pct:.1f}%)" if impact_amt>0 else '-',
                '归因分析(主因)':reason_str,
                '转化率变化':f"{cvr_diff:+.1f}pp" if cvr_diff is not None else '--',
                '流量变化':_pct(vis_chg_m)})
        st.dataframe(df(tbl_rows),use_container_width=True,hide_index=True,height=max(280,min(520,len(tbl_rows)*38+40)))
        top_drag=sorted(ch_model_issues,key=lambda x:x['环比'])[:5]
        drag_summary=' | '.join([f"[{t['型号']}]{_pct(t['环比'])}" for t in top_drag])
        st.caption(f'\U0001F4CC 最大拖累TOP5: {drag_summary}')
    else:
        st.info('\u2705 所有渠道的各型号表现稳定，未发现显著异常下滑。')

    # --- C2. 转化率骤降型号 ---
    st.markdown('#### \U0001F4C9 转化率骤降型号（降幅>20%且访客>50）')
    cvr_drop_models=[]
    for mk_key,mv in cur_by_model.items():
        pv_m=prev_by_model.get(mk_key,{})
        cvr_c=mv.get('支付转化率',0);cvr_p=pv_m.get('支付转化率',0)
        if cvr_p>=0.005 and cvr_c<cvr_p:
            cvr_drop=(cvr_c-cvr_p)/cvr_p
            if cvr_drop<-0.20 and mv.get('商品访客数',0)>50:
                cvr_drop_models.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '本期转化率':cvr_c*100,'上期转化率':cvr_p*100,'降幅':cvr_drop,
                    '本期访客':mv.get('商品访客数',0),'本期GMV':mv.get('支付金额',0),
                    '本期加购人数':mv.get('商品加购人数',0),'上期加购人数':pv_m.get('商品加购人数',0)})
    cvr_drop_models.sort(key=lambda x:x['降幅'])
    if cvr_drop_models:
        cdr=[]
        for c in cvr_drop_models[:15]:
            cart_c=c.get('本期加购人数',0); cart_p=c.get('上期加购人数',0)
            vis_c=c.get('本期访客',0)
            fn=''
            if cart_c and cart_p and vis_c:
                crc=cart_c/vis_c*100; crp=cart_p/vis_c*100
                if crc<crp-2: fn=f'(加购率也\u2193{crc-crp:.1f}pp\u2192详情页吸引力下降)'
                else: fn=f'(加购率正常\u2192可能为价格/评价/库存因素)'
            cdr.append({'渠道':c['渠道'],'品类':c['品类'],'型号':c['型号'],
                '上期转化':f"{c['上期转化率']:.2f}%",'本期转化':f"{c['本期转化率']:.2f}%",
                '降幅':f"<span style='color:#dc2626;font-weight:700'>{_pct(c['降幅'])}</span>",
                '本期访客':f"{c['本期访客']:,.0f}",'本期GMV':f"\u00A5{c['本期GMV']:,.0f}",
                '漏斗判断':fn})
        st.dataframe(df(cdr),use_container_width=True,hide_index=True,height=min(450,len(cdr)*36+40))
    else:
        st.info('\u2705 未发现转化率骤降型号（阈值：降幅>20%，访客>50）。')

    # --- C3. 客单价明显下跌型号 ---
    st.markdown('#### \U0001F4B0 客单价明显下跌型号（降幅超10%）')
    aov_drop=[]
    for mk_key,mv in cur_by_model.items():
        pv_m=prev_by_model.get(mk_key,{})
        ac=mv.get('客单价',0);ap=pv_m.get('客单价',0)
        if ap>10 and ac<ap:
            ad=(ac-ap)/ap
            if ad<-0.10 and mv.get('支付件数',0)>10:
                aov_drop.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '上期客单':ap,'本期客单':ac,'降幅':ad,
                    '本期件数':mv.get('支付件数',0),'上期件数':pv_m.get('支付件数',0),
                    '本期GMV':mv.get('支付金额',0)})
    aov_drop.sort(key=lambda x:x['降幅'])
    if aov_drop:
        adr=[]
        for a in aov_drop[:15]:
            pc_c=a.get('本期件数',0);pc_p=a.get('上期件数',0)
            ar=''
            if pc_c>pc_p*1.3: ar='(件数\u2191但均价\u2193\u2192低价SKU占比提升/折扣加大)'
            elif pc_c<pc_p*0.7: ar='(件数\u2193且均价\u2193\u2192高客单SKU销量萎缩)'
            else: ar='(件数持平\u2192直接降价/促销力度加大)'
            adr.append({'渠道':a['渠道'],'品类':a['品类'],'型号':a['型号'],
                '上期客单价':f"\u00A5{a['上期客单']:,.0f}",'本期客单价':f"\u00A5{a['本期客单']:,.0f}",
                '降幅':f"<span style='color:#ea580c;font-weight:700'>{_pct(a['降幅'])}</span>",
                '本期件数':f"{a['本期件数']:,.0f}",'上期件数':f"{a['上期件数']:,.0f}",
                '初步判断':ar})
        st.dataframe(df(adr),use_container_width=True,hide_index=True,height=min(420,len(adr)*36+40))
    else:
        st.info('\u2705 未发现客单价明显下跌型号（阈值：降幅>10%，件数>10）。')

    # --- C4. 爆款断崖掉量 ---
    st.markdown('#### \u26A1 爆款断崖式掉量型号（上期TOP20\u2192本期缩水>30%）')
    prev_top20=sorted(prev_by_model.items(),key=lambda x:x[1].get('支付金额',0),reverse=True)[:20]
    drop_stars=[]
    for mk_key,pv in prev_top20:
        mv=cur_by_model.get(mk_key,{})
        pc=mv.get('支付金额',0);pp=pv.get('支付金额',0)
        if pp>0:
            drop=(pc-pp)/pp
            if drop<-0.30:
                drop_stars.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '上期GMV':pp,'本期GMV':pc,'缩水幅度':drop})
    if drop_stars:
        dsr=[]
        for d in drop_stars[:15]:
            loss=d['上期GMV']-d['本期GMV']
            dsr.append({'渠道':d['渠道'],'品类':d['品类'],'型号':d['型号'],
                '上期GMV':f"\u00A5{d['上期GMV']:,.0f}",'本期GMV':f"\u00A5{d['本期GMV']:,.0f}",
                '缩水':f"<span style='color:#dc2626;font-weight:700'>{_pct(d['缩水幅度'])}</span>",
                '损失金额':f"\u00A5{loss:,.0f}",
                '占上期份额':f"{d['上期GMV']/max(prev_sum_all.get('支付金额',1),1)*100:.1f}%"})
        st.dataframe(df(dsr),use_container_width=True,hide_index=True,height=min(400,len(dsr)*36+40))
    else:
        st.info('\u2705 上期TOP20爆款型号均保持稳定。')

    # --- C5. 新晋增长亮点（新增正向反馈）---
    st.markdown('#### \U0001F31F 新晋增长亮点（增速>50%或新上榜）')
    cur_top20=sorted(cur_by_model.items(),key=lambda x:x[1].get('支付金额',0),reverse=True)[:20]
    rising_stars=[]
    for mk_key,mv in cur_top20:
        pv_m=prev_by_model.get(mk_key,{})
        mc=mv.get('支付金额',0);mp=pv_m.get('支付金额',0)
        if mp>0:
            growth=(mc-mp)/mp
            if growth>0.50 and mc>1000:
                rising_stars.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '上期GMV':mp,'本期GMV':mc,'增速':growth})
        else:
            if mc>2000:
                rising_stars.append({'渠道':mk_key[0],'品类':mk_key[1],'型号':mk_key[2],
                    '上期GMV':0,'本期GMV':mc,'增速':float('inf')})
    rising_stars.sort(key=lambda x:x['本期GMV'],reverse=True)
    if rising_stars:
        rsr=[]
        for r in rising_stars[:10]:
            sp="<span style='color:#22c55e;font-weight:700'>\U0001F680 新品/爆发</span>" if r['增速']==float('inf') else f"<span style='color:#22c55e;font-weight:700'>+{r['增速']*100:.0f}%</span>"
            rsr.append({'渠道':r['渠道'],'品类':r['品类'],'型号':r['型号'],
                '上期GMV':f"\u00A5{r['上期GMV']:,.0f}" if r['上期GMV']>0 else '新上榜',
                '本期GMV':f"\u00A5{r['本期GMV']:,.0f}",'增速':sp})
        st.dataframe(df(rsr),use_container_width=True,hide_index=True,height=min(340,len(rsr)*34+40))
    else:
        st.info('\u2139\ufe0f 本周期未发现显著增长新星（阈值：增速>50% 或 新上榜且GMV>\u00A52000）。')

    # ══════════════════════════════════════
    # D. 第三层：具体可执行的优化措施（绑定真实数据值）
    # ══════════════════════════════════════
    st.markdown('<hr style="margin:18px 0;border:none;border-top:1px dashed #cbd5e1;">')
    st.markdown('<div class="section-title">\U0001F527 具体执行措施清单</div>',unsafe_allow_html=True)
    st.caption(f"以下根据上述诊断结果自动生成（共识别 {len(ch_model_issues)} 个异常型号）。每条措施绑定实际数据值。")

    actions=[]

    def add_action(priority,title,detail,owner,timeline,metric_target):
        actions.append({'p':priority,'t':title,'d':detail,'o':owner,'tl':timeline,'mt':metric_target})

    # ======== GMV 问题 ========
    if gmv_g is not None and gmv_g<0:
        gmv_loss=max(0, prev_sum_all.get('支付金额',0)-cur_sum.get('支付金额',0))
        if vis_g is not None and vis_g<-0.08:
            vis_loss_share = abs(vis_g)/(abs(vis_g)+abs(cvr_g or 0)+abs(aov_g or 0)+0.001)
            add_action('P0','【流量】紧急排查核心渠道流量断崖',
                f'<b>现状：</b>访客从 {prev_sum_all.get("商品访客数",0):,.0f}\u2192{cur_sum.get("商品访客数",0):,.0f}（{_pct(vis_g)}），'
                f'潜在影响GMV约 \u00A5{gmv_loss*vis_loss_share:,.0f}<br><br>'
                f'<b>排查步骤：</b><br>'
                f'\u2460 直通车后台\u2192推广计划列表\u2192筛选近7天展现量\u219330%\u2192检查预算耗尽/质量分下降<br>'
                f'\u2461 生意参谋\u2192市场\u2192搜索分析\u2192核心类目Top10词\u2192对比搜索人气和点击率<br>'
                f'\u2462 直播间：查看回放\u2192流量来源分析\u2192确认付费/免费比例<br>'
                f'\u2463 <b>应急：</b>表现最差的3个计划日预算+50%，观察3天',
                '运营负责人','24小时内',f'访客\u2265{prev_sum_all.get("商品访客数",0)*0.95:,.0f}')
        if cvr_g is not None and cvr_g<-0.08:
            ccvr=cur_sum.get('支付转化率',0)*100; pcvr=prev_sum_all.get('支付转化率',0)*100
            lost_orders=cur_sum.get('商品访客数',0)*(prev_sum_all.get('支付转化率',0)-cur_sum.get('支付转化率',0))
            lost_gmv_val=lost_orders*cur_sum.get('客单价",0)
            add_action('P0','【转化】全店转化率紧急提升行动',
                f'<b>现状：</b>{pcvr:.2f}%\u2192{ccvr:.2f}%（\u2193{pcvr-ccvr:.2f}pp），'
                f'少成交约{lost_orders:,.0f}单，影响\u00A5{lost_gmv_val:,.0f}<br><br>'
                f'<b>执行步骤：</b><br>'
                f'\u2460 从「转化率骤降表」提取全部异常SKU\u2192逐一做首屏3秒测试<br>'
                f'\u2461 导出近90天评价\u2192词频统计\u2192负面Top3\u2192修改FAQ和卖点<br>'
                f'\u2462 平台搜索这些型号\u2192对比竞品前3名\u2192高于竞品8%则设限时9折<br>'
                f'\u2463 检查大促是否刚结束\u2192价格回调导致骤降\u2192延长优惠3天',
                '运营+美工','3天内',f'转化率\u2265{pcvr*0.97:.2f}%')
        if aov_g is not None and aov_g<-0.06:
            add_action('P1','【客单价】高客单价SKU曝光恢复',
                f'<b>现状：</b>\u00A5{prev_sum_all.get("客单价",0):.0f}\u2192\u00A5{cur_sum.get("客单价",0):.0f}（{_pct(aov_g)}）<br><br>'
                f'\u2460 提取客单价前20SKU\u2192核对本周vs上周访客\u2192圈出降幅最大的5个<br>'
                f'\u2461 设置关联推荐：「搭配购买减X」「买二送一」，放在加购区下方<br>'
                f'\u2462 满减门槛：均值\u00A5{cur_sum.get("客单价",0):.0f}\u2192满减线设\u00A5{cur_sum.get("客单价",0)*1.3:,.0f}<br>'
                f'\u2463 直通车\u2192「高消费力」人群溢价+20%',
                '运营','1周内',f'客单价\u2265\u00A5{prev_sum_all.get("客单价",0)*0.97:,.0f}')

    # ======== 渠道级别措施 ========
    if ch_model_issues:
        ch_gmv_changes={}
        for ck,cv_ch in cur_by_channel.items():
            pv_ch=prev_by_channel.get(ck,{})
            cc=cv_ch.get('支付金额',0);pp=pv_ch.get('支付金额',0)
            ch_gmv_changes[ck[0]]=(cc-pp)/pp if pp else None
        sorted_ch=sorted(ch_gmv_changes.items(),key=lambda x:x[1] if x[1] else 0)
        worst_ch=sorted_ch[0] if sorted_ch else (None,None)
        if worst_ch[1] and worst_ch[1]<-0.05:
            ch_nm=worst_ch[0];ch_pct=_pct(worst_ch[1])
            bad_in=[m for m in ch_model_issues if m['渠道']==ch_nm][:3]
            ml=', '.join([f"[{m['型号']}]({m['品类']})" for m in bad_in]) or '\u591A\u4e2a\u578b\u53f7'
            add_action('P0',f'【渠道】{ch_nm}专项整改（GMV{ch_pct}）',
                f'该渠道GMV{ch_pct}，集中在：{ml}<br>'
                f'\u2460 检查上述型号推广状态（\u505c/\u964d\u6743/\u8FDD\u89C4）<br>'
                f'\u2461 DSR\u5206\u6570\uff1aDSR&lt;4.7\u5F71\u54CD\u641C\u7D22\u6743\u91CD<br>'
                f'\u2462 \u6838\u5BF9\u6D3B\u52A8\u62A5\u540D\u60C5\u51B5\uff0C\u91CD\u8981\u4F1A\u573A\u8865\u62A5<br>'
                f'\u2463 \u6296\u97F3\u6E20\u9053\uFF1A\u67E5\u8FD17\u5929\u76F4\u64AD\u65F6\u957F\u548CGMV/\u5C0F\u65F6',
                f'{ch_nm}\u6E20\u9053\u8D1F\u8D23\u4EBA','48\u5C0F\u65F6\u5185',f'{ch_nm} GMV\u73AF\u6BD4\u8F6C\u6B63')

    # ======== 型号级别具体措施（Top3异常型号逐一给方案）=====
    if ch_model_issues:
        top3_bad=ch_model_issues[:3]
        for bm in top3_bad:
            mod=bm['型号'];cat=bm['品类'];ch=bm['渠道']
            chg_pct=_pct(bm['环比'])
            vc=bm['本期访客'];vp=bm['上期访客']
            vm=(vc-vp)/vp if vp else None
            ccr=bm['本期转化率']*100;cpr=bm['上期转化率']*100
            cdr=ccr-cpr
            root='\u6D41\u91CF\u65AD\u5D16' if (vm and vm<-0.15) else ('\u8F6C\u5316\u5931\u6548' if cdr<-2 else '\u590D\u5408\u8870\u9000')
            pri='P0' if bm['环比']<DANGER_T else 'P1'
            loss_amnt=bm['上期GMV']-bm['本期GMV']
            detail=(
                f'<b>数据：</b>\u00A5{bm["\u672C\u671FGMV"]:,.0f} vs \u00A5{bm["\u4E0A\u671FGMV"]:,.0f}'
                f'(\u635F\u5931\u00A5{loss_amnt:,.0f})\uff0c\u8F6C\u5316{ccr:.2f}% vs {cpr:.2f}%<br><br>'
                f'<b>[{root}] \u5B9A\u5411\u65BD\uFF1A</b><br>')
            if root=='\u6D41\u91CF\u65AD\u5D16':
                detail+=(
                    f'\u2460 <b>\u6D41\u91CF\u7AEF</b>\uFF1A\u68C0\u67E5\u8BE5\u578B\u53F7\u5728{ch}\u6E20\u9053\u7684\u641C\u7D22\u6392\u540D\u3001\u4E3B\u56FECTR\uff0C'
                    f'CTR={vc/(vp+0.001)*100:.1f}%({"<3%\u9700\u6362\u4E3B\u56FE" if vp and vc/vp<0.03 else "\u6B63\u5E38"})<br>'
                    f'\u2461 \u68C0\u67E5\u662F\u5426\u6709\u63A8\u5E7F\u8BA1\u5212\u88AB\u7CFB\u7EDF\u9650\u6D41/\u9884\u7B97\u8017\u5C3D<br>'
                    f'\u2462 <b>\u6025\u6551</b>\uFF1A\u4E34\u65F6\u589E\u52A0\u76F4\u901A\u8F66\u65E5\u9884\u7B97+50%\uFF0C\u6301\u7EED3\u5929\u89C2\u5BDF<br>')
            else:
                detail+=(
                    f'\u2460 <b>\u8F6C\u5316\u7AEF</b>\uFF1A\u6253\u5F00\u8BE5\u578B\u53F7\u8BE6\u60C5\u9875\u6A21\u62DF\u4E70\u5BB6\u6D4F\u89C8\u2014\u2014\u9996\u5C4F3\u79D2\u80FD\u5426\u770B\u6E05\u6838\u5FC3\u5356\u70B9\uFF1F<br>'
                    f'\u2461 <b>\u8BC4\u4EF7\u5BA1\u8BA1</b>\uFF1A\u5BFC\u51FA\u8FD160\u5929\u8BC4\u4EF7\u2192\u8BCD\u9891\u7EDF\u8BA1\u2192\u8D1F\u9762Top3\u2192\u4F18\u5316\u8BDD\u672F<br>'
                    f'\u2462 <b>\u4EF7\u683C\u5BF9\u6807</b>\uFF1A\u641C\u7D22\u540C\u6B3E\u7ADE\u54C13\u5BB6\uFF0C'
                    f'{"\u9AD8\u4E8E\u7ADE\u54C18\u5219\u8BBE\u9650\u65F6\u6298" if bm["\u672C\u671F\u5BA2\u5355\u4EF7"] > 0 and bm["\u4E0A\u671F\u5BA2\u5355\u4EF7"] > 0 and bm["\u672C\u671F\u5BA2\u5355\u4EF7"] >= bm["\u4E0A\u671F\u5BA2\u5355\u4EF7"]*1.08 else "\u4EF7\u683C\u57FA\u672C\u5408\u7406"}<br>'
                    f'\u2463 <b>\u5E93\u5B58\u68C0\u67E5</b>\uFF1A\u786E\u8BA4\u8BE5\u578B\u53F7\u65E0\u7F3A\u8D27/\u9884\u552E\u72B6\u6001<br>'
                    f'\u2464 <b>\u5DEE\u8BC4\u5904\u7406</b>\uFF1A\u7B5B\u9009\u51FA\u73B0\u22652\u6B21\u4EE5\u4E0A\u7684\u8D1F\u9762\u6807\u7B7E\u96C6\u4E2D\u5904\u7406')
            add_action(pri,f'[{mod}]({cat}/{ch}){root}-GMV{chg_pct}',
                detail,f'\u8FD0\u8425-{cat}\u7EC4','3-5\u5929\u89C1\u6548',f'{mod} GMV\u73AF\u6BD4\u2191>-5%')

    # ======== 退款率措施 ========
    if ref_g is not None and ref_g>0.05:
        crp=cur_sum.get('退款率',0)*100
        if crp>8 or ref_g>0.10:
            add_action('P1','【\u552E\u540E】\u9000\u6B3E\u7387\u5F02\u5E38\u5347\u9AD8',
                f'\u5F53\u524D{crp:.1f}%\uff0C\u53D8\u5316{_pct(ref_g)}<br>'
                f'\u2460 \u5BFC\u51FA\u8FD130\u5929\u9000\u6B3E\u8BA2\u5355\u6309\u300C\u9000\u6B3E\u539F\u56E0\u300D\u5E01\u5E15\u62C9\u5206析\uFF0CTop3\u539F\u56E0<br>'
                f'\u2461 「\u8D28\u91CF\u95EE\u9898/\u63CF\u8FF0\u4E0D\u7B26」&gt;40%\uFF1A\u8D28\u68C0\u56E2\u961F\u62BD\u68C0<br>'
                f'\u2462 「\u7269\u6D41\u6162/\u7834\u635F」&gt;30%\uFF1A\u6539\u8FDB\u5305\u88C5+\u5207\u6362\u5FEB\u9012<br>'
                f'\u2463 「\u4E0D\u60F3\u8981\u4E86\"\uFF1C\u8BF4\u660E\u8BE6\u60C5\u9875\u8BEF\u5BFC\u4FE1\u606F\uFF0C\u9700\u4FEE\u6B63',
                '\u5BA2\u670D+\u4ED3\u50A8+\u8D28\u68C0','2\u5468\u5185','\u9000\u6B3E\u7387&lt;5%')

    # ======== 常规措施 ========
    add_action('P3','【\u5E38\u89C4】\u6BCF\u5468\u4E00\u4E0A\u5348\u5065\u5EB7\u68C0\u67E5',
        '\u6BCF\u5468\u4E00 10:00\u5B8C\u6210：<br>'
        '\u2460 \u6253\u5F00\u672C\u770B\u677F\u300C\u667A\u80FD\u8BCA\u65AD\u300DTab\u622A\u56FE\u5B58\u6863<br>'
        '\u2461 \u5BF9\u6BD4\u4E0A\u5468\u540C\u671F\u6807\u8BB0\u53D8\u5316\u00B1\u5%<br>'
        '\u2462 \u8FDE\u7EED2\u5468\u540C\u4E00\u6307\u6807\u4E0B\u6ed1\u2192\u4E13\u9879\u4F1A\u8BAE<br>'
        '\u2463 \u68C0\u67E5\u672C\u5468\u5230\u671F\u6D3B\u52A8/\u4F18\u60E0\u5238\u7EED\u671F',
        '\u8FD0\u8425\u8D1F\u8D23\u4EBA','\u6BCF\u5468\u4E00\u56FA\u5B9A','\u5468\u62A5\u5B58\u6863')
    add_action('P3','【\u5E38\u89C4】\u6708\u5EA6\u6E20\u9053ROI\u590D\u76D8',
        '\u6BCF\u67085\u65E5\u524D\u5B8C\u6210\u4E0A\u6708\u5404\u6E20\u9053ROI\uFF1A<br>'
        f'ROI=(\u6E20\u9053\u9500\u552E\u989D-\u9000\u8D27\u989D)/\u6E20\u9053\u63A8\u5E7F\u8D39\u7528<br>'
        '\u2460 &gt;5 \u52A0\u5927\u6295\u5165 / 2-5 \u7EF4\u6301 / &lt;2 \u7F29\u51CF\u6216\u4F18\u5316<br>'
        '\u2461 ROI&lt;2\u8F93\u51FA\u300AXX\u6E20\u9053\u4F18\u5316\u65B9\u6848\u300B<br>',
        '\u8FD0\u8425+\u8D22\u52A1','\u6BCF\u67085\u53F7\u524D','\u5168\u6E20\u9053\u5747ROI&gt;3')

    # ---- 展示 ----
    actions_sorted=sorted(actions,key=lambda x:['P0','P1','P2','P3'].index(x['p']))
    for act in actions_sorted:
        cls={'P0':'tag-p0','P1':'tag-p1','P2':'tag-p2','P3':'tag-p3'}[act['p']]
        tag_html=f"<span class='action-tag {cls}'>{act['p']}</span>"
        exp_title=f"{tag_html} **{act['t']}** <small style='color:#94a3b8;'>| {act['o']} | \u76EE\u6807: {act['mt']} | \u89C1\u6548: {act['tl']}</small>"
        with st.expander(exp_title, expanded=(act['p']=='P0')):
            st.markdown(act['d'], unsafe_allow_html=True)
    if not actions_sorted:
        st.success('\u2705 \u5F53\u524D\u6240\u6709\u6838\u5FC3\u6307\u6807\u5065\u5EB7\uFF0C\u65E0\u989D\u5916\u5E72\u9884\u3002')

    # ══════════════════════════════════════
    # E. 一键下载诊断报告（修复列错乱问题）
    # ══════════════════════════════════════
    st.markdown("<hr style='margin:18px 0;border:none;border-top:1px dashed #cbd5e1;'>")
    rep_header=['\u8BCA\u65AD\u65F6\u95F4','\u8BCA\u65AD\u533A\u95F4','\u5BF9\u6BD4\u533A\u95F4','GMV\u53D8\u5316','\u8BBF\u5BA2\u53D8\u5316',
               '\u8F6C\u5316\u7387\u53D8\u5316','\u5BA2\u5355\u4EF7\u53D8\u5316','\u9000\u6B3E\u7387','\u5065\u5EB7\u8BC4\u5206',
               '\u5F02\u5E38\u578B\u53F7\u6570','\u8F6C\u5316\u9AA4\u964D\u578B\u53F7\u6570','\u7206\u6B3E\u6389\u91CF\u6570','\u65B0\u661F\u589E\u957F\u6570',
               'P0\u4EFB\u52A1\u6570','P1\u4EFB\u52A1\u6570']
    dl_data=[{
        '诊断时间': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), '诊断区间': f'{s}~{e}', '对比区间': f'{prev_s_d}~{prev_e_d}',
        'GMV变化': _pct(gmv_g), '访客变化': _pct(vis_g),
        '转化率变化': f"{(cur_sum.get('支付转化率',0)-prev_sum_all.get('支付转化率',0))*100:+.2f}pp" if prev_sum_all.get('支付转化率',0) else '--',
        '客单价变化': _pct(aov_g), '退款率变化': _pct(ref_g),
        '健康评分': f'{health_score:.0f}/100', '健康结论': hv[0],
        '异常型号数': len(ch_model_issues), '转化骤降数': len(cvr_drop_models),
        '爆款掉量数': len(drop_stars), '新星增长数': len(rising_stars) if 'rising_stars' in dir() else 0,
        'P0任务数': sum(1 for a in actions if a['p']=='P0'), 'P1任务数': sum(1 for a in actions if a['p']=='P1'),
    }]
    for act in actions_sorted:
        dl_data.append({
            '诊断时间': '', '诊断区间': '', '对比区间': '',
            'GMV变化': '', '访客变化': '', '转化率变化': '', '客单价变化': '', '退款率变化': '',
            '健康评分': '', '健康结论': '',
            '异常型号数': '', '转化骤降数': '', '爆款掉量数': '', '新星增长数': '',
            'P0任务数': '', 'P1任务数': '',
            '优先级': act['p'], '措施标题': act['t'], '负责人': act['o'], '见效周期': act['tl'], '量化目标': act['mt'],
        })
    st.download_button('\U0001F4E5 \u4E0B\u8F7D\u5B8C\u6574\u8BCA\u65AD\u62A5\u544A CSV',
        rows_to_csv(dl_data, list(dl_data[0].keys()) if dl_data else rep_header),
        file_name=f'xiaotunbi_diagnosis_{s.replace("-","")}_{e.replace("-","")}.csv',
        mime='text/csv')
'''

# 定位旧Tab 6开始位置
marker = "# TAB 6: 智能诊断"
idx = content.find(marker)
if idx == -1:
    print("ERROR: Cannot find Tab 6 marker!")
else:
    print(f"Found Tab 6 at position {idx}")
    # 替换从Tab 6到文件末尾的所有内容
    new_content = content[:idx] + new_tab6
    
    # 写回文件
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    original_len = len(content[idx:])
    new_len = len(new_tab6)
    print(f"Replacement done! Original: {original_len} chars -> New: {new_len} chars")
    print(f"Total file size: {len(new_content)} chars")
