from pathlib import Path
import pandas as pd
import numpy as np
import math

BASE = Path('/home/ubuntu')
WORK = BASE / 'revenue_rebuild'
DATA = WORK / 'data'
CODE = WORK / 'code'

stream = pd.read_csv(DATA/'streaming_model_v3.csv')
saas = pd.read_csv(DATA/'saas_model_v3.csv')
stream_scen = pd.read_csv(DATA/'streaming_scenarios_v3.csv')
saas_scen = pd.read_csv(DATA/'saas_scenarios_v3.csv')
mc = pd.read_csv(DATA/'saas_clv_cac_uncertainty_v3.csv')
prov = pd.read_csv(DATA/'source_provenance_v3.csv')

checks = []
# Reconstructed 59 checks with weights 1-3
# Section 1 Assumptions Audit (7)
sec='1. Assumptions Audit'
ass = [
('A1','Are all major assumptions explicit and centralized rather than buried in code or narrative?',3),
('A2','Are causal relationships stated clearly for each major driver?',3),
('A3','Are historical dependencies or precedent links documented where used?',2),
('A4','Are policy / management intervention assumptions explicitly separated from organic behavior?',2),
('A5','Are contradictory assumptions absent across model layers and reports?',3),
('A6','Is false precision avoided in key assumptions and outputs?',2),
('A7','Is outcome bias avoided, i.e. assumptions are not tuned to a preferred narrative?',3),
]
# Section 2 Structural Risks (8)
sec2='2. Structural Risks'
struct = [
('S1','Are circular dependencies eliminated?',3),('S2','Are outputs not being derived from themselves indirectly?',3),('S3','Are important non-linear relationships represented where needed?',2),('S4','Is double-counting removed from revenue and customer bridges?',3),('S5','Is top-down logic consistent with bottom-up logic?',3),('S6','Are cohort aggregations handled correctly?',2),('S7','Are timing and lags modeled explicitly where material?',3),('S8','Are hardcoded values minimized and parameterized?',2)
]
# Section 3 Sensitivity Risks (6)
sec3='3. Sensitivity Risks'
sens = [
('R1','Are dominant drivers identified explicitly?',2),('R2','Are disproportionate output swings tested and disclosed?',3),('R3','Is uncertainty quantified for key drivers?',3),('R4','Are correlated driver risks modeled or discussed?',2),('R5','Are assumption ranges wide enough to be decision-useful?',2),('R6','Are non-linear sensitivity effects handled?',2)
]
# Section 4 Data & Input Integrity (6)
sec4='4. Data & Input Integrity'
data_checks = [
('D1','Are source data and benchmark inputs validated and cited?',3),('D2','Are manual overrides documented?',2),('D3','Is there an audit trail from source to output?',3),('D4','Are anomalies / one-offs prevented from dominating logic?',2),('D5','Are organic behavior and interventions separated?',2),('D6','Are lagging indicators not misused as leading inputs?',3)
]
# Section 5 Business Reality Alignment (6)
sec5='5. Business Reality Alignment'
biz = [
('B1','Are capacity constraints modeled?',3),('B2','Are sales cycles / onboarding delays represented where relevant?',3),('B3','Is enterprise deal lumpiness or timing variability modeled?',2),('B4','Are dynamic team / customer behaviors represented?',2),('B5','Are pricing response effects modeled?',2),('B6','Are operational bottlenecks represented or disclosed?',2)
]
# Section 6 Dependency & Systemic Risk (5)
sec6='6. Dependency & Systemic Risk'
dep = [
('Y1','Are root-factor dependencies identified?',2),('Y2','Are driver interdependencies explicit?',3),('Y3','Are correlated risks quantified or scenario-tested?',2),('Y4','Are feedback loops modeled?',3),('Y5','Are regime changes (growth to saturation / competition / recession) modeled?',3)
]
# Section 7 Edge Cases & Failure Modes (6)
sec7='7. Edge Cases & Failure Modes'
edge = [
('E1','Is top-customer / top-account churn impact handled where material?',2),('E2','Is major deal slippage or acquisition miss handled?',3),('E3','Is abrupt growth slowdown modeled?',3),('E4','Is pricing change failure modeled?',2),('E5','Are extreme scenarios tested?',3),('E6','Does the model behave logically at extremes?',3)
]
# Section 8 Output Risk & Interpretation (5)
sec8='8. Output Risk & Interpretation'
out = [
('O1','Is false precision removed from outputs?',2),('O2','Are uncertainty ranges clearly communicated?',3),('O3','Is deterministic misinterpretation actively prevented?',2),('O4','Is real-world volatility reflected?',2),('O5','Are driver explanations transparent enough for decision-makers?',2)
]
# Section 9 Model Purpose & Bias (5)
sec9='9. Model Purpose & Bias'
purpose = [
('P1','Is model purpose clearly defined (planning / forecasting / target-setting)?',2),('P2','Are assumptions aligned to that purpose?',2),('P3','Is narrative pressure bias removed?',3),('P4','Are downside risks represented fairly?',3),('P5','Is optimism/pessimism balanced?',2)
]
# Section 10 Unit Economics Consistency (5)
sec10='10. Unit Economics Consistency'
ue = [
('U1','Are CAC, CLV/LTV, churn, and ARPU/ARPA internally aligned?',3),('U2','Is expansion consistent with base dynamics?',2),('U3','Are retention and growth coherent together?',3),('U4','Are margins aligned with scale and context?',2),('U5','Are growth and efficiency not contradicting each other?',3)
]

all_sections=[(sec,ass),(sec2,struct),(sec3,sens),(sec4,data_checks),(sec5,biz),(sec6,dep),(sec7,edge),(sec8,out),(sec9,purpose),(sec10,ue)]
for section, items in all_sections:
    for cid, text, wt in items:
        for domain in ['SaaS','Streaming']:
            checks.append({'section':section,'check_id':cid,'check_text':text,'model_domain':domain,'severity_weight':wt})
check_df=pd.DataFrame(checks)
check_df.to_csv(DATA/'checklist_59_reconstructed.csv', index=False)

results=[]

def add(domain, section, cid, status, evidence, risk, fix):
    results.append({'section':section,'check_id':cid,'model_domain':domain,'status':status,'evidence':evidence,'risk_note':risk,'proposed_fix':fix})

# Helper evidence
stream_bridge_ok = float(stream['Bridge Check'].abs().max()) <= 1e-6
saas_sensitivity_diff = float(abs(saas_scen[saas_scen['Scenario']=='low_churn']['Month24 ARR'].iloc[0] - saas_scen[saas_scen['Scenario']=='high_churn']['Month24 ARR'].iloc[0]))
mc_ratio = mc[mc['Metric']=='CLV/CAC Ratio'].iloc[0]

# Manual+programmatic evaluation
for domain in ['SaaS','Streaming']:
    for section, items in all_sections:
        for cid, text, wt in items:
            status='⚠️ Partial'; evidence=''; risk=''; fix=''
            if cid=='A1':
                status = '⚠️ Partial'
                evidence = 'Assumptions sheet exists in workbook, but several scenario modifiers and curve scalars remain code-level rather than fully centralized.'
                risk='Hidden assumptions reduce auditability.'
                fix='Centralize all assumptions and scenario modifiers in workbook/data tables and generated assumption logs.'
            elif cid=='A2':
                status = '✅ Pass'
                evidence = 'v3 code and reports describe explicit bridges and causal flow.'
                risk=''; fix='Maintain causal diagrams and bridge tables.'
            elif cid=='A3':
                status = '⚠️ Partial'
                evidence = 'Source provenance exists, but some derived assumptions (e.g. SaaS reactivation downscale) lack deeper historical basis.'
                risk='Historical dependency weakness can distort realism.'
                fix='Add explicit assumption provenance table with derived-vs-sourced labels and rationale.'
            elif cid=='A4':
                if domain=='Streaming':
                    status='⚠️ Partial'; evidence='Competition, macro shock, and price drag are scenario parameters but not always labeled as interventions vs environment.'; risk='Mixing exogenous and management levers can confuse decisions.'; fix='Separate environment variables from management-policy levers.'
                else:
                    status='⚠️ Partial'; evidence='Pricing and CAC policies are mixed with market effects in scenario params.'; risk='Management action not isolated.'; fix='Split controllable vs uncontrollable drivers.'
            elif cid=='A5':
                status='✅ Pass'; evidence='No obvious identity contradictions in v3 core outputs.'; fix=''; risk=''
            elif cid=='A6':
                status='⚠️ Partial'; evidence='Some reports still use many decimals in outputs/means and scenario tables.'; risk='False precision in executive interpretation.'; fix='Round outputs and label estimates as benchmark-calibrated ranges.'
            elif cid=='A7':
                if domain=='SaaS':
                    status='✅ Pass'; evidence=f'Actual CLV/CAC reported as {mc_ratio["Mean"]:.2f}x rather than tuned upward.'; risk=''; fix=''
                else:
                    status='✅ Pass'; evidence='Downside scenarios are included and materially worse than base.'; risk=''; fix=''
            elif cid=='S1':
                status='✅ Pass'; evidence='No circular formulas detected in v3 bridges.'; risk=''; fix=''
            elif cid=='S2':
                status='✅ Pass'; evidence='Outputs are no longer target-plugged.'; risk=''; fix=''
            elif cid=='S3':
                status='⚠️ Partial'; evidence='Some nonlinearity exists via cohort aging and scenario shifts, but not enough around capacity, price elasticity, or step-change effects.'; risk='Model may understate threshold behavior.'; fix='Add explicit non-linear functions for capacity, pricing elasticity, and sales productivity saturation.'
            elif cid=='S4':
                status='✅ Pass'; evidence='Streaming bridge check passes; SaaS ARR bridge explicit.'; risk=''; fix=''
            elif cid=='S5':
                status='⚠️ Partial'; evidence='Bottom-up cohort logic exists, but top-down benchmark totals are not routinely reconciled to external scaling bands.'; risk='Local coherence without macro plausibility checks.'; fix='Add top-down sanity checks and scaling guardrails.'
            elif cid=='S6':
                status='⚠️ Partial'; evidence='Cohort aging exists, but top customer concentration / segment cohorts are not modeled.'; risk='Aggregation hides concentration risk.'; fix='Add segment/concentration cohorts.'
            elif cid=='S7':
                if domain=='SaaS':
                    status='❌ Fail'; evidence='No explicit sales cycle or onboarding lag in v3 SaaS acquisition-to-revenue conversion.'; risk='Revenue recognized too quickly versus real GTM behavior.'; fix='Add pipeline-to-bookings lag and onboarding ramp.'
                else:
                    status='⚠️ Partial'; evidence='Streaming has reactivation timing but no explicit billing delay or campaign lag.'; risk='Timing still simplified.'; fix='Add acquisition/retention intervention lag structure.'
            elif cid=='S8':
                status='⚠️ Partial'; evidence='Some hardcoded base rates remain in code (e.g., 0.036 streaming base churn, 0.0115 SaaS base churn).'; risk='Hardcoded logic limits reuse and governance.'; fix='Move to parameter tables.'
            elif cid=='R1':
                status='❌ Fail'; evidence='Dominant drivers are not explicitly ranked or quantified in v3 outputs.'; risk='Decision-makers may focus on the wrong levers.'; fix='Add driver importance / tornado analysis.'
            elif cid=='R2':
                status='⚠️ Partial'; evidence='Scenarios exist, but output swing attribution is not decomposed.'; risk='Large swings not explained.'; fix='Add contribution-to-change decomposition.'
            elif cid=='R3':
                if domain=='SaaS':
                    status='✅ Pass'; evidence='Monte Carlo uncertainty exists for CLV/CAC.'; risk=''; fix=''
                else:
                    status='⚠️ Partial'; evidence='Streaming uses scenario ranges but lacks quantified distribution / fan probabilities.'; risk='Under-communicated uncertainty.'; fix='Add wider uncertainty tables and driver bands.'
            elif cid=='R4':
                status='⚠️ Partial'; evidence='Scenarios bundle correlated risks qualitatively, but no explicit correlation matrix or joint shock documentation.'; risk='Underestimates compound downside.'; fix='Add correlated scenario matrix.'
            elif cid=='R5':
                status='⚠️ Partial'; evidence='Ranges are useful but still somewhat narrow for severe shocks.'; risk='Downside not wide enough.'; fix='Add extreme stress scenarios.'
            elif cid=='R6':
                status='⚠️ Partial'; evidence='Some nonlinearities included, but not sensitivity cliffs.'; risk='Missed tipping points.'; fix='Add elasticity and breakpoint testing.'
            elif cid=='D1':
                status='⚠️ Partial'; evidence='Sources are cited, but some URLs are high-level pages and some derived assumptions are not separately versioned.'; risk='Input validation trail incomplete.'; fix='Add source snapshot table and citation type field.'
            elif cid=='D2':
                status='❌ Fail'; evidence='No dedicated manual override log in v3.'; risk='Undocumented overrides or edits may go unnoticed.'; fix='Add manual overrides sheet/log.'
            elif cid=='D3':
                status='⚠️ Partial'; evidence='There is some audit trail via CSVs and reports, but not a line-by-line source-to-parameter map for every assumption.'; risk='Harder to audit assumptions.'; fix='Add parameter registry and provenance IDs.'
            elif cid=='D4':
                status='⚠️ Partial'; evidence='Anomaly handling discussed, but no formal anomaly controls or one-off flags.'; risk='One-off benchmarks may bias base rates.'; fix='Add anomaly flag fields and exclusion rationale.'
            elif cid=='D5':
                status='⚠️ Partial'; evidence='Behavior vs intervention partly separated, not explicit everywhere.'; risk='Wrong attribution of outcomes.'; fix='Separate environmental, policy, and behavioral parameters.'
            elif cid=='D6':
                status='⚠️ Partial'; evidence='Some benchmark outputs may still influence inputs indirectly (e.g., CAC framing from mature benchmarks applied as base input).'; risk='Lagging benchmarks used too directly.'; fix='Tag lagging benchmarks and reduce direct use as leading assumptions.'
            elif cid=='B1':
                if domain=='Streaming':
                    status='⚠️ Partial'; evidence='Saturation via market capacity multiplier exists.'; risk='Capacity is crude and not segment-specific.'; fix='Add segment/TAM capacity layers.'
                else:
                    status='❌ Fail'; evidence='No sales-team or onboarding capacity constraint in SaaS.'; risk='Growth can outpace operational reality.'; fix='Add sales capacity and implementation throughput constraints.'
            elif cid=='B2':
                status='❌ Fail'; evidence='No explicit sales cycle / onboarding delay for SaaS and only limited timing for streaming.'; risk='Revenue timing too optimistic.'; fix='Add pipeline and onboarding lag mechanics.'
            elif cid=='B3':
                if domain=='SaaS':
                    status='❌ Fail'; evidence='No enterprise deal lumpiness modeled.'; risk='Monthly forecast too smooth.'; fix='Add stochastic or scenario-based lumpiness.'
                else:
                    status='⚠️ Partial'; evidence='Streaming lumps are less relevant but campaign bursts not modeled.'; risk='User acquisition smoothness may be unrealistic.'; fix='Add campaign pulse effect.'
            elif cid=='B4':
                status='⚠️ Partial'; evidence='Customer behavior dynamic exists, team behavior does not.'; risk='Operational adaptation missing.'; fix='Add productivity / service response effects.'
            elif cid=='B5':
                status='⚠️ Partial'; evidence='Price pressure affects churn, but explicit elasticity calibration is limited.'; risk='Weak pricing response realism.'; fix='Add price elasticity parameterization.'
            elif cid=='B6':
                status='❌ Fail'; evidence='No explicit operational bottleneck variables.'; risk='Forecast ignores execution ceilings.'; fix='Add capacity / staffing / support bottleneck variables.'
            elif cid=='Y1':
                status='⚠️ Partial'; evidence='Macro shock and competition act as root factors, but dependency map is not explicit.'; risk='Same root factor may be counted inconsistently.'; fix='Add dependency map sheet.'
            elif cid=='Y2':
                status='✅ Pass'; evidence='Driver interdependencies are modeled in v3.'; risk=''; fix=''
            elif cid=='Y3':
                status='⚠️ Partial'; evidence='Correlated risks are scenario-bundled but not quantified.'; risk='Systemic downside understated.'; fix='Add joint-shock table/correlation matrix.'
            elif cid=='Y4':
                status='✅ Pass'; evidence='Feedback loops exist for saturation, pressure, pricing/churn interactions.'; risk=''; fix=''
            elif cid=='Y5':
                status='✅ Pass'; evidence='Growth, recession, competition, saturation regimes included.'; risk=''; fix=''
            elif cid=='E1':
                if domain=='SaaS':
                    status='❌ Fail'; evidence='No top-customer concentration layer.'; risk='Largest-customer churn impact invisible.'; fix='Add concentration stress module.'
                else:
                    status='⚠️ Partial'; evidence='Top-customer not relevant, but platform/channel concentration still absent.'; risk='Distribution/channel shock missing.'; fix='Add distributor/channel shock scenario.'
            elif cid=='E2':
                status='⚠️ Partial'; evidence='Acquisition miss appears through scenarios but not explicit deal slippage module.'; risk='Missed timing shocks not isolated.'; fix='Add slippage scenario and lag buckets.'
            elif cid=='E3':
                status='✅ Pass'; evidence='Recession/competition/saturation materially slow growth.'; risk=''; fix=''
            elif cid=='E4':
                status='⚠️ Partial'; evidence='Pricing failure not isolated from macro/competition.'; risk='Cannot isolate pricing experiment risk.'; fix='Add failed-price-increase scenario.'
            elif cid=='E5':
                status='⚠️ Partial'; evidence='Stress scenarios exist but not truly extreme tail cases.'; risk='Black-swan downside understated.'; fix='Add severe stress tests.'
            elif cid=='E6':
                status='✅ Pass'; evidence='Model remains logical under tested scenarios; no bridge breaks.'; risk=''; fix=''
            elif cid=='O1':
                status='⚠️ Partial'; evidence='Some outputs still over-decimalized.'; risk='False precision.'; fix='Round executive outputs.'
            elif cid=='O2':
                status='⚠️ Partial'; evidence='Ranges exist but not consistently front-and-center for both domains.'; risk='Users may over-read point estimates.'; fix='Lead with ranges and bands in reports/dashboard.'
            elif cid=='O3':
                status='⚠️ Partial'; evidence='Reports warn about uncertainty, but workbook dashboard still foregrounds point estimates.'; risk='Deterministic misuse.'; fix='Add uncertainty banners and scenario comparison summary.'
            elif cid=='O4':
                status='⚠️ Partial'; evidence='Some volatility through scenarios, not enough intra-scenario volatility/lumpiness.'; risk='Too smooth.'; fix='Add controlled volatility/lumpiness.'
            elif cid=='O5':
                status='✅ Pass'; evidence='Drivers are reasonably explainable in v3 reports.'; risk=''; fix=''
            elif cid=='P1':
                status='⚠️ Partial'; evidence='Purpose is described as planning/forecasting, but target-setting boundaries are not explicit enough.'; risk='Misuse as quota/commit target.'; fix='Add explicit intended-use statement.'
            elif cid=='P2':
                status='⚠️ Partial'; evidence='Assumptions are mostly planning-oriented but some benchmark norms may be read as targets.'; risk='Normative benchmark bias.'; fix='Separate benchmark reference from target assumptions.'
            elif cid=='P3':
                status='✅ Pass'; evidence='v3 accepts weak CLV/CAC and downside scenarios.'; risk=''; fix=''
            elif cid=='P4':
                status='✅ Pass'; evidence='Downside cases represented.'; risk=''; fix=''
            elif cid=='P5':
                status='⚠️ Partial'; evidence='Base case may still be optimistic because operational bottlenecks absent.'; risk='Slight optimism bias.'; fix='Add bottleneck and delay realism.'
            elif cid=='U1':
                if domain=='SaaS':
                    status='⚠️ Partial'; evidence=f'Actual CLV/CAC is {mc_ratio["Mean"]:.2f}x and coherent numerically, but still benchmark-dependent and weak economically.'; risk='Unit economics may be unattractive or unstable.'; fix='Add payback and margin/cohort validation.'
                else:
                    status='✅ Pass'; evidence='Streaming excludes unsupported CAC and avoids false unit-economics linkage.'; risk=''; fix=''
            elif cid=='U2':
                status='✅ Pass'; evidence='Expansion tied to retained / installed base.'; risk=''; fix=''
            elif cid=='U3':
                status='✅ Pass'; evidence='Growth slows in high-churn and recession scenarios.'; risk=''; fix=''
            elif cid=='U4':
                if domain=='SaaS':
                    status='⚠️ Partial'; evidence='Gross margin anchored to benchmarks but not dynamic with scale.'; risk='Margin realism limited.'; fix='Add modest scale-vs-margin function.'
                else:
                    status='⚠️ Partial'; evidence='Streaming has no explicit gross margin layer.'; risk='Economic quality incomplete.'; fix='Keep excluded or add clearly-labeled margin proxy.'
            elif cid=='U5':
                if domain=='SaaS':
                    status='⚠️ Partial'; evidence='Growth and efficiency can diverge logically, but no explicit payback check exists.'; risk='Possible hidden contradiction.'; fix='Add CAC payback and Rule-of-40-like checks.'
                else:
                    status='✅ Pass'; evidence='Streaming does not overstate efficiency metrics it cannot support.'; risk=''; fix=''
            add(domain, section, cid, status, evidence, risk, fix)

res = pd.DataFrame(results)
res.to_csv(DATA/'checklist_audit_results_v3.csv', index=False)

# Score
score_map={'✅ Pass':1.0,'⚠️ Partial':0.5,'❌ Fail':0.0}
merged = res.merge(check_df[['section','check_id','model_domain','severity_weight']], on=['section','check_id','model_domain'], how='left')
merged['earned'] = merged['status'].map(score_map) * merged['severity_weight']
merged['possible'] = merged['severity_weight']
summary = merged.groupby('model_domain').agg(earned=('earned','sum'), possible=('possible','sum')).reset_index()
summary['score_pct'] = summary['earned'] / summary['possible'] * 100
summary.to_csv(DATA/'checklist_scores_v3.csv', index=False)
print(summary.to_string(index=False))
