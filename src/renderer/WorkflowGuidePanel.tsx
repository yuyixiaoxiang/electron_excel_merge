import React, { useMemo, useState } from 'react';

type WorkflowMode = 'diff' | 'merge';
type GuideTab = 'overview' | 'origins' | 'capabilities' | 'workflows';

type WorkflowStep = {
  title: string;
  description: string;
};

export interface WorkflowGuidePanelProps {
  mode: WorkflowMode;
  cliMode?: 'diff' | 'merge' | null;
  autoHasPrimaryKey?: boolean;
  primaryKeyHint?: string;
  showFullTables?: boolean;
  frozenRowCount?: number;
  frozenColCount?: number;
  mergeFrozenRowCount?: number;
  rowSimilarityThreshold?: number;
}

const tabs: Array<{ key: GuideTab; label: string; hint: string }> = [
  { key: 'overview', label: '总览', hint: '当前模式与运行摘要' },
  { key: 'origins', label: '设计来源', hint: 'jsdiff / ExcelMerge 借鉴点' },
  { key: 'capabilities', label: '关键能力', hint: '主键与公式边界' },
  { key: 'workflows', label: '模式工作流', hint: 'diff / merge 两条路径' },
];

const containerStyle: React.CSSProperties = {
  marginBottom: 12,
  border: '1px solid #d6dfee',
  borderRadius: 12,
  backgroundColor: '#fcfdff',
  overflow: 'hidden',
  flexShrink: 0,
  boxShadow: '0 1px 2px rgba(15, 23, 42, 0.04)',
};

const headerStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  gap: 12,
  padding: '12px 14px',
  borderBottom: '1px solid #e7edf6',
  backgroundColor: '#f6f8fc',
};

const headerTitleStyle: React.CSSProperties = {
  margin: 0,
  fontSize: 14,
  fontWeight: 700,
  color: '#1f2937',
};

const headerSubtitleStyle: React.CSSProperties = {
  marginTop: 4,
  fontSize: 12,
  color: '#5f6b7a',
  lineHeight: 1.5,
};

const actionButtonStyle: React.CSSProperties = {
  border: '1px solid #c8d2e3',
  backgroundColor: '#fff',
  borderRadius: 8,
  padding: '5px 12px',
  cursor: 'pointer',
  fontSize: 12,
  color: '#32425c',
  flexShrink: 0,
};

const heroStyle: React.CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: 12,
  alignItems: 'stretch',
  padding: 14,
  border: '1px solid #d7e6ff',
  borderRadius: 12,
  background: 'linear-gradient(135deg, #f8fbff 0%, #f4f8ff 58%, #ffffff 100%)',
};

const heroBadgeStyle: React.CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  gap: 6,
  padding: '3px 9px',
  borderRadius: 999,
  fontSize: 11,
  fontWeight: 600,
  color: '#2155a3',
  border: '1px solid #cfe0fb',
  backgroundColor: '#ffffff',
};

const heroTitleStyle: React.CSSProperties = {
  margin: '10px 0 0 0',
  fontSize: 18,
  fontWeight: 700,
  color: '#142033',
  lineHeight: 1.35,
};

const heroDescriptionStyle: React.CSSProperties = {
  marginTop: 8,
  fontSize: 12,
  color: '#526273',
  lineHeight: 1.7,
};

const infoCardStyle: React.CSSProperties = {
  flex: '1 1 320px',
  minWidth: 280,
  border: '1px solid #e4ebf5',
  borderRadius: 12,
  backgroundColor: '#fff',
  padding: 12,
  boxSizing: 'border-box',
};

const emphasisCardStyle: React.CSSProperties = {
  ...infoCardStyle,
  borderColor: '#d7e6ff',
  backgroundColor: '#fbfdff',
};

const tabBarStyle: React.CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: 8,
};

const tabButtonStyle = (active: boolean): React.CSSProperties => ({
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'flex-start',
  gap: 4,
  minWidth: 140,
  borderRadius: 10,
  border: `1px solid ${active ? '#c5dafc' : '#d9e1ec'}`,
  backgroundColor: active ? '#eef5ff' : '#fff',
  color: active ? '#13438c' : '#344255',
  padding: '8px 10px',
  cursor: 'pointer',
  fontSize: 12,
  textAlign: 'left',
});

const chipStyle = (active: boolean): React.CSSProperties => ({
  display: 'inline-flex',
  alignItems: 'center',
  padding: '3px 9px',
  borderRadius: 999,
  fontSize: 12,
  fontWeight: 500,
  border: `1px solid ${active ? '#97c0ff' : '#d4d9e2'}`,
  backgroundColor: active ? '#e8f2ff' : '#fff',
  color: active ? '#0f4aa1' : '#556070',
});

const cardTitleStyle: React.CSSProperties = {
  fontSize: 13,
  fontWeight: 700,
  color: '#172236',
};

const cardDescriptionStyle: React.CSSProperties = {
  marginTop: 6,
  fontSize: 12,
  color: '#5a6675',
  lineHeight: 1.65,
};

const listStyle: React.CSSProperties = {
  margin: '8px 0 0 18px',
  padding: 0,
  fontSize: 12,
  color: '#2c3848',
  lineHeight: 1.7,
};

const contentWrapStyle: React.CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: 12,
};

const metricBoxStyle: React.CSSProperties = {
  borderRadius: 10,
  border: '1px solid #dce6f5',
  backgroundColor: '#fff',
  padding: '10px 12px',
};

const metricLabelStyle: React.CSSProperties = {
  fontSize: 11,
  color: '#66758a',
};

const metricValueStyle: React.CSSProperties = {
  marginTop: 4,
  fontSize: 12,
  fontWeight: 600,
  color: '#19263c',
  lineHeight: 1.5,
};

const miniModeCardStyle = (active: boolean): React.CSSProperties => ({
  borderRadius: 10,
  border: `1px solid ${active ? '#b9d6ff' : '#dce2ea'}`,
  backgroundColor: active ? '#f3f8ff' : '#fff',
  padding: '10px 12px',
});

const workflowCardStyle = (active: boolean): React.CSSProperties => ({
  ...infoCardStyle,
  flex: '1 1 420px',
  borderColor: active ? '#9ec7ff' : '#dfe6f0',
  boxShadow: active ? '0 0 0 1px rgba(37, 99, 235, 0.08)' : 'none',
  backgroundColor: active ? '#fcfdff' : '#fff',
});

const stepCardStyle = (active: boolean): React.CSSProperties => ({
  display: 'flex',
  gap: 10,
  alignItems: 'flex-start',
  padding: '9px 10px',
  borderRadius: 10,
  border: `1px solid ${active ? '#d2e5ff' : '#e3e8ef'}`,
  backgroundColor: active ? '#f8fbff' : '#fafbfc',
});

const stepIndexStyle = (active: boolean): React.CSSProperties => ({
  width: 22,
  height: 22,
  borderRadius: 999,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  fontSize: 11,
  fontWeight: 700,
  color: active ? '#0b4aa0' : '#5c6778',
  backgroundColor: active ? '#dcebff' : '#edf1f5',
  flexShrink: 0,
  marginTop: 1,
});

const modeLabel = (mode: WorkflowMode) => (mode === 'diff' ? '双文件 diff 模式' : '三方 merge 模式');

const modeSummaryLabel = (mode: WorkflowMode, cliMode?: 'diff' | 'merge' | null) => {
  if (mode === 'diff') return '默认主流程是左右两个 Excel 的并排对比，可直接编辑并分别保存。';
  if (cliMode === 'merge') return '当前由 Git/Fork merge 驱动，重点是决策后回写 MERGED/ours。';
  if (cliMode === 'diff') return '当前由 Git/Fork diff 驱动，重点是审查 ours 与 theirs 的差异。';
  return '当前为交互式三方 merge，重点是主键对齐、差异决策和预览回写。';
};

export const WorkflowGuidePanel: React.FC<WorkflowGuidePanelProps> = ({
  mode,
  cliMode,
  autoHasPrimaryKey = true,
  primaryKeyHint,
  showFullTables = false,
  frozenRowCount = 0,
  frozenColCount = 0,
  mergeFrozenRowCount = 3,
  rowSimilarityThreshold = 0.9,
}) => {
  const [expanded, setExpanded] = useState(true);
  const [activeTab, setActiveTab] = useState<GuideTab>('overview');

  const primaryKeyStatus = useMemo(() => {
    if (primaryKeyHint) return primaryKeyHint;
    if (mode === 'diff') {
      return autoHasPrimaryKey ? '当前 diff 按主键列稳定行身份。' : '当前 diff 退化为无主键的序列/内容对齐。';
    }
    return autoHasPrimaryKey ? '当前按主键列做行身份对齐。' : '当前退化为无主键的序列/内容对齐。';
  }, [mode, autoHasPrimaryKey, primaryKeyHint]);

  const runtimeSummary = useMemo(() => {
    if (mode === 'diff') {
      return `当前为默认双文件 diff 路径：视图冻结行=${mergeFrozenRowCount}，行相似度阈值=${rowSimilarityThreshold.toFixed(2)}。左右两栏可以直接编辑并分别保存。`;
    }
    const cliSummary =
      cliMode === 'merge'
        ? '当前由 Git/Fork merge 模式驱动，保存会写回 MERGED/ours。'
        : cliMode === 'diff'
          ? '当前由 Git/Fork diff 模式驱动，保存会覆盖 ours。'
          : '当前为交互式 merge 模式，保存时由用户选择目标文件。';
    const tableSummary = showFullTables ? '已开启 ours/theirs 全表查看。' : '当前只聚焦差异行列。';
    return `${cliSummary} 视图冻结行=${mergeFrozenRowCount}，行相似度阈值=${rowSimilarityThreshold.toFixed(2)}。${tableSummary}`;
  }, [mode, cliMode, showFullTables, mergeFrozenRowCount, rowSimilarityThreshold]);

  const modeGuide = useMemo(() => {
    if (mode === 'diff') {
      return {
        title: 'diff：默认双文件左右对比',
        summary: '适合把两个 Excel 放到左右两栏里直接审查和编辑，像 Beyond Compare 一样以“比较”为主、以“分别保存”为辅。',
        focus: '重点关注左右文件选择、并排编辑、对齐后的差异高亮，以及分别写回 left / right 文件。',
      };
    }
    return {
      title: 'merge：围绕 base / ours / theirs 做决策',
      summary: '适合处理真实合并问题：先做行身份对齐，再逐格、逐行、逐列选择，最后统一从 merged preview 落盘。',
      focus: modeSummaryLabel(mode, cliMode),
    };
  }, [mode, cliMode]);

  const quickMetrics = useMemo(
    () => [
      { label: '当前模式', value: modeLabel(mode) },
      {
        label: '主键策略',
        value: autoHasPrimaryKey ? '优先按主键对齐' : '退化为序列/内容对齐',
      },
      { label: '公式语义', value: '按结果比较，尽量保留未触碰公式' },
    ],
    [mode, autoHasPrimaryKey],
  );

  const workflowGuides = useMemo<Record<WorkflowMode, { subtitle: string; summary: string; footer: string; steps: WorkflowStep[] }>>(
    () => ({
      diff: {
        subtitle: '适合“左文件 / 右文件 → 对齐 → 审查 → 就地编辑 → 分别保存”的默认路径',
        summary: 'diff 模式不再单独强调“打开一个文件”，而是把双文件并排对比作为主产品形态。',
        footer: `当前 diff 运行参数：视图冻结行=${mergeFrozenRowCount}，行相似度阈值=${rowSimilarityThreshold.toFixed(2)}。`,
        steps: [
          { title: '选择 left / right 两个 Excel', description: '底部固定文件选择器决定左右两栏的文件来源，非 Git merge 启动时默认就走这条路径。' },
          { title: '按对齐结果并排审查', description: '工作表会按同名 sheet 对齐，左右表格同步滚动，适合像 Beyond Compare 一样扫差异。' },
          { title: '双击单元格直接编辑', description: '已有对应单元格的位置可以直接修改，修改会暂存在当前侧，便于边比边改。' },
          { title: '分别保存 left / right', description: '左右文件各自维护未保存修改，确认后分别写回原始 Excel。' },
        ],
      },
      merge: {
        subtitle: '适合“base / ours / theirs → 对齐 → 决策 → 统一回写”的合并路径',
        summary: 'merge 模式把三方差异、主键身份、结构操作和 merged preview 放进同一条决策链路里。',
        footer:
          cliMode === 'merge'
            ? '当前 merge 由 Git/Fork merge 触发，保存会回写 MERGED/ours。'
            : cliMode === 'diff'
              ? '当前 merge 由 Git/Fork diff 触发，保存会覆盖 ours。'
              : '当前 merge 为交互式打开，保存时由用户选择目标文件。',
        steps: [
          { title: '打开 base / ours / theirs', description: '按工作表名与索引对齐后进入 side-by-side diff，形成三方比较上下文。' },
          { title: '调整主键 / 视图冻结 / 阈值', description: '用主键列定义行身份；冻结行只影响审查视角，不改变 diff/merge 计算；相似度阈值才会影响对齐结果。' },
          { title: '逐格、逐行、逐列做选择', description: '可以选 ours / theirs，也可以生成插删行列操作，把结构变化一并纳入决策。' },
          { title: '从 merged preview 统一写回', description: '下方预览区实时拼出最终结果，确认后再统一保存到目标文件。' },
        ],
      },
    }),
    [cliMode, mergeFrozenRowCount, rowSimilarityThreshold],
  );

  const renderTabContent = () => {
    if (activeTab === 'overview') {
      return (
        <div style={contentWrapStyle}>
          <div style={{ ...infoCardStyle, flex: '2 1 500px' }}>
            <div style={cardTitleStyle}>这块面板现在承担什么角色</div>
            <div style={cardDescriptionStyle}>
              这里不再是简单备注，而是把产品的设计来源、关键边界和两种工作模式的操作路径，直接嵌进界面里。
            </div>
            <ul style={listStyle}>
              <li>先回答“当前模式是什么、适合做什么、运行参数意味着什么”。</li>
              <li>再解释“这个产品借鉴了谁、为什么需要主键、公式到底能做到哪一步”。</li>
              <li>最后把 diff 与 merge 两条工作流拆开，让用户在同一屏内切换理解。</li>
            </ul>
          </div>
          <div style={emphasisCardStyle}>
            <div style={cardTitleStyle}>当前模式建议</div>
            <div style={{ marginTop: 8, fontSize: 14, fontWeight: 700, color: '#133d7a' }}>{modeGuide.title}</div>
            <div style={cardDescriptionStyle}>{modeGuide.summary}</div>
            <div
              style={{
                marginTop: 10,
                padding: '9px 10px',
                borderRadius: 10,
                backgroundColor: '#fff',
                border: '1px solid #dae7fb',
                fontSize: 12,
                color: '#2d435d',
                lineHeight: 1.65,
              }}
            >
              {modeGuide.focus}
            </div>
          </div>
          <div style={infoCardStyle}>
            <div style={cardTitleStyle}>运行时要点</div>
            <div style={{ marginTop: 8, display: 'grid', gap: 8 }}>
              {quickMetrics.map((item) => (
                <div key={item.label} style={metricBoxStyle}>
                  <div style={metricLabelStyle}>{item.label}</div>
                  <div style={metricValueStyle}>{item.value}</div>
                </div>
              ))}
            </div>
          </div>
          <div style={infoCardStyle}>
            <div style={cardTitleStyle}>切换模式时记住</div>
            <div style={{ marginTop: 8, display: 'grid', gap: 8 }}>
              <div style={miniModeCardStyle(mode === 'diff')}>
                <div style={{ fontSize: 12, fontWeight: 700, color: '#1d293c' }}>diff</div>
                <div style={cardDescriptionStyle}>
                  偏“比较器”视角：关注 left / right 两栏、差异高亮、就地编辑和分别保存。
                </div>
              </div>
              <div style={miniModeCardStyle(mode === 'merge')}>
                <div style={{ fontSize: 12, fontWeight: 700, color: '#1d293c' }}>merge</div>
                <div style={cardDescriptionStyle}>
                  偏“决策器”视角：关注 base / ours / theirs、主键对齐、冲突选择与 merged preview。
                </div>
              </div>
            </div>
          </div>
        </div>
      );
    }

    if (activeTab === 'origins') {
      return (
        <div style={contentWrapStyle}>
          <div style={infoCardStyle}>
            <div style={cardTitleStyle}>借鉴自 jsdiff</div>
            <div style={cardDescriptionStyle}>偏算法与工程化：让 diff 逻辑可以拆层、替换、测试和限流。</div>
            <ul style={listStyle}>
              <li>把 diff 拆成分词 / 比较器 / 变化对象输出，方便后续替换策略与调试。</li>
              <li>强调 timeout、maxEditLength、可中断与边界测试，适合大表 diff 的性能兜底。</li>
              <li>给当前项目的启发是：重复行、空值、阈值、中止条件都应该被算法层显式建模。</li>
            </ul>
          </div>
          <div style={infoCardStyle}>
            <div style={cardTitleStyle}>借鉴自 ExcelMerge</div>
            <div style={cardDescriptionStyle}>偏 Excel 专用体验：用表格结构和人工导航能力提升审查效率。</div>
            <ul style={listStyle}>
              <li>强调 `Sheet / Row / Cell Diff` 的 Excel 专用数据模型，而不是只盯一张扁平矩阵。</li>
              <li>把 row header / column header / 搜索 / 跳转 / 汇总做成高频操作，服务人工核对。</li>
              <li>给当前项目的启发是：主键、视觉坐标映射和导航信息也应该成为 UI 的一部分。</li>
            </ul>
          </div>
          <div style={{ ...emphasisCardStyle, flex: '1 1 100%' }}>
            <div style={cardTitleStyle}>落地到当前产品后的表达</div>
            <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 8 }}>
              <div style={{ ...metricBoxStyle, flex: '1 1 220px' }}>
                <div style={metricLabelStyle}>算法层</div>
                <div style={metricValueStyle}>向 jsdiff 学模块化与性能兜底，不把所有语义塞进 UI。</div>
              </div>
              <div style={{ ...metricBoxStyle, flex: '1 1 220px' }}>
                <div style={metricLabelStyle}>数据模型层</div>
                <div style={metricValueStyle}>向 ExcelMerge 学 sheet / row / cell 语义和视觉坐标映射。</div>
              </div>
              <div style={{ ...metricBoxStyle, flex: '1 1 220px' }}>
                <div style={metricLabelStyle}>产品层</div>
                <div style={metricValueStyle}>把主键、公式边界、模式差异做成界面内文档，而不是口头说明。</div>
              </div>
            </div>
          </div>
        </div>
      );
    }

    if (activeTab === 'capabilities') {
      return (
        <div style={contentWrapStyle}>
          <div style={{ ...infoCardStyle, borderColor: '#d6e6ff', backgroundColor: '#fbfdff' }}>
            <div style={cardTitleStyle}>主键：行身份系统，不是普通筛选条件</div>
            <div style={cardDescriptionStyle}>
              主键决定“哪一行和哪一行是同一条业务记录”。一旦这个身份判错，后面的插删行、改单元格、冲突判断都会被带偏。
            </div>
            <ul style={listStyle}>
              <li>有主键时优先按 key 对齐；无主键时才退化到序列/内容相似度对齐。</li>
              <li>既要支持手动指定，也要支持自动识别和列对齐后的逻辑列映射。</li>
              <li>它的价值是稳定行身份，而不是单纯“找一个比较像 ID 的列”。</li>
            </ul>
            <div
              style={{
                marginTop: 10,
                padding: '8px 10px',
                borderRadius: 10,
                backgroundColor: '#fff',
                border: '1px solid #d8e6fb',
                fontSize: 12,
                color: '#194b93',
                lineHeight: 1.65,
              }}
            >
              当前状态：{primaryKeyStatus}
            </div>
          </div>
          <div style={{ ...infoCardStyle, borderColor: '#f0c7bd', backgroundColor: '#fffaf9' }}>
            <div style={cardTitleStyle}>公式：当前是“结果导向”的支持，不是公式语义合并</div>
            <div style={cardDescriptionStyle}>
              现在更擅长比较公式算出来的结果，并尽量保住未触碰的公式模板；但公式表达式本身还不是一等公民。
            </div>
            <ul style={listStyle}>
              <li>当前 diff 主要比较公式结果值，而不是公式表达式本体。</li>
              <li>未触碰的公式单元格通常能随 `ours` 模板保留下来；直接改写时会退化成常量值。</li>
              <li>结构操作之后，普通引用与 shared formula 仍然有风险，需要单独关注。</li>
            </ul>
            <div
              style={{
                marginTop: 10,
                padding: '8px 10px',
                borderRadius: 10,
                backgroundColor: '#fff',
                border: '1px solid #f1d4ce',
                fontSize: 12,
                color: '#8d3b2c',
                lineHeight: 1.65,
              }}
            >
              适合“结果导向”的合并；如果公式本身就是业务逻辑，仍需要更强的 raw-cell / formula-aware 模型。
            </div>
          </div>
          <div style={{ ...emphasisCardStyle, flex: '1 1 100%' }}>
            <div style={cardTitleStyle}>什么时候可以放心用，什么时候要提高警惕</div>
            <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 12 }}>
              <div style={{ ...metricBoxStyle, flex: '1 1 260px' }}>
                <div style={metricLabelStyle}>更适合</div>
                <div style={metricValueStyle}>有稳定主键、公式主要作为结果展示、用户需要高效人工审查与合并决策。</div>
              </div>
              <div style={{ ...metricBoxStyle, flex: '1 1 260px' }}>
                <div style={metricLabelStyle}>需要谨慎</div>
                <div style={metricValueStyle}>公式本体就是业务逻辑、共享公式很多、或者频繁插删行列改变引用关系。</div>
              </div>
            </div>
          </div>
        </div>
      );
    }

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div
          style={{
            borderRadius: 10,
            border: '1px solid #d7e6ff',
            backgroundColor: '#f7fbff',
            padding: '10px 12px',
            fontSize: 12,
            color: '#33506e',
            lineHeight: 1.65,
          }}
        >
          当前高亮的是 <strong>{modeLabel(mode)}</strong>；另一栏则保留完整路径，方便你在 diff / merge 之间快速切换上下文。
        </div>
        <div style={contentWrapStyle}>
          {(['diff', 'merge'] as WorkflowMode[]).map((targetMode) => {
            const guide = workflowGuides[targetMode];
            const active = mode === targetMode;
            return (
              <div key={targetMode} style={workflowCardStyle(active)}>
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, alignItems: 'flex-start' }}>
                  <div>
                    <div style={cardTitleStyle}>
                      {targetMode === 'diff' ? '工作模式一：diff（双文件左右对比）' : '工作模式二：merge（base / ours / theirs）'}
                    </div>
                    <div style={cardDescriptionStyle}>{guide.subtitle}</div>
                  </div>
                  {active && <span style={chipStyle(true)}>当前激活</span>}
                </div>
                <div style={{ marginTop: 10, fontSize: 12, color: '#415166', lineHeight: 1.65 }}>{guide.summary}</div>
                <div style={{ marginTop: 12, display: 'grid', gap: 8 }}>
                  {guide.steps.map((step, index) => (
                    <div key={`${targetMode}-${step.title}`} style={stepCardStyle(active)}>
                      <span style={stepIndexStyle(active)}>{index + 1}</span>
                      <div style={{ minWidth: 0 }}>
                        <div style={{ fontSize: 12, fontWeight: 700, color: '#1b283b' }}>{step.title}</div>
                        <div style={cardDescriptionStyle}>{step.description}</div>
                      </div>
                    </div>
                  ))}
                </div>
                <div
                  style={{
                    marginTop: 12,
                    borderRadius: 10,
                    border: '1px solid #e1e8f1',
                    backgroundColor: '#fff',
                    padding: '9px 10px',
                    fontSize: 12,
                    color: '#526173',
                    lineHeight: 1.65,
                  }}
                >
                  {guide.footer}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  return (
    <div style={containerStyle}>
      <div style={headerStyle}>
        <div style={{ minWidth: 0 }}>
          <h3 style={headerTitleStyle}>工作流 / 设计说明</h3>
          <div style={headerSubtitleStyle}>
            把当前 UI 的设计来源、主键 / 公式边界，以及 diff / merge 两条工作路径，整理成一块内置产品文档。
          </div>
        </div>
        <button type="button" onClick={() => setExpanded((prev) => !prev)} style={actionButtonStyle}>
          {expanded ? '收起说明' : '展开说明'}
        </button>
      </div>
      {expanded && (
        <div style={{ padding: 14, display: 'flex', flexDirection: 'column', gap: 14 }}>
          <div style={heroStyle}>
            <div style={{ flex: '2 1 500px', minWidth: 320 }}>
              <span style={heroBadgeStyle}>内置工作流指南</span>
              <div style={heroTitleStyle}>把“为什么这么设计”和“现在该怎么用”放到同一屏里。</div>
              <div style={heroDescriptionStyle}>
                这块面板不只是说明文字，而是把当前模式、借鉴来源、主键与公式边界、以及 diff / merge 两种工作流整理成随时可切换的界面内文档。
              </div>
              <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
                <span style={chipStyle(true)}>当前模式：{modeLabel(mode)}</span>
                <span style={chipStyle(mode === 'diff')}>diff</span>
                <span style={chipStyle(mode === 'merge')}>merge</span>
                {mode === 'merge' && cliMode && <span style={chipStyle(true)}>Git 调用：{cliMode}</span>}
              </div>
            </div>
            <div style={{ ...emphasisCardStyle, flex: '1 1 280px', minWidth: 280 }}>
              <div style={cardTitleStyle}>当前运行摘要</div>
              <div style={cardDescriptionStyle}>{runtimeSummary}</div>
              <div
                style={{
                  marginTop: 10,
                  padding: '8px 10px',
                  borderRadius: 10,
                  backgroundColor: '#fff',
                  border: '1px solid #dde8f7',
                  fontSize: 12,
                  color: '#24415f',
                  lineHeight: 1.65,
                }}
              >
                {modeSummaryLabel(mode, cliMode)}
              </div>
            </div>
          </div>

          <div style={tabBarStyle}>
            {tabs.map((tab) => {
              const active = tab.key === activeTab;
              return (
                <button key={tab.key} type="button" onClick={() => setActiveTab(tab.key)} style={tabButtonStyle(active)}>
                  <span style={{ fontSize: 12, fontWeight: 700 }}>{tab.label}</span>
                  <span style={{ fontSize: 11, color: active ? '#3d5c88' : '#708093' }}>{tab.hint}</span>
                </button>
              );
            })}
          </div>

          {renderTabContent()}
        </div>
      )}
    </div>
  );
};
