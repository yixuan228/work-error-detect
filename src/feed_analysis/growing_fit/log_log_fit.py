
"""
Docstring for feed_analysis.growing_fit.log_log_fit

1) 包括训练模型；可视化模型拟合结果
2
"""


import numpy as np
import statsmodels.api as sm
from sklearn.preprocessing import PolynomialFeatures, StandardScaler

def poly_log_regression_with_smearing(
    df,
    x_col="age",
    y_col="avg_food_kg",
    degree=2,
    alpha=0.15,
    show_metrics=True
):
    """
    对 y 做 log1p，对 x 做 log 变换 + 多项式回归，
    使用 smearing 修正还原到原始尺度，并给出置信区间与预测区间。

    Parameters
    ----------
    df : pd.DataFrame
        输入数据
    x_col : str
        自变量列名（如 age）
    y_col : str
        因变量列名（如 avg_food_kg）
    degree : int
        多项式阶数（默认 2）
    alpha : float
        置信水平（默认 0.15 → 85% CI）
    show_metrics : bool
        是否打印评估指标（R² / RMSE / MAE 等）

    Returns
    -------
    df_out : pd.DataFrame
        含预测值和区间的新 DataFrame
    model : statsmodels.regression.linear_model.RegressionResults
        拟合好的模型
    """

    df = df.copy()

    # --------------------
    # 构造 X / y
    # --------------------
    X = df[[x_col]].copy()
    X[x_col] = np.log(X[x_col] + 1e-6)
    y = df[y_col]
    y_log = np.log1p(y)

    # --------------------
    # 多项式特征
    # --------------------
    poly = PolynomialFeatures(degree=degree, include_bias=False)
    X_poly = poly.fit_transform(X)

    # 标准化
    scaler = StandardScaler()
    X_poly_scaled = scaler.fit_transform(X_poly)

    # --------------------
    # statsmodels 回归
    # --------------------
    X_poly_sm = sm.add_constant(X_poly_scaled)
    model = sm.OLS(y_log, X_poly_sm).fit()

    # --------------------
    # 预测 + 区间
    # --------------------
    pred = model.get_prediction(X_poly_sm)
    pred_summary = pred.summary_frame(alpha=alpha)

    # smearing 修正
    smearing = np.mean(np.exp(model.resid))

    df["food_fit"] = smearing * np.expm1(pred_summary["mean"])
    df["ci_lower"] = smearing * np.expm1(pred_summary["mean_ci_lower"])
    df["ci_upper"] = smearing * np.expm1(pred_summary["mean_ci_upper"])

    df["pred_lower_food"] = smearing * np.expm1(pred_summary["obs_ci_lower"])
    df["pred_upper_food"] = smearing * np.expm1(pred_summary["obs_ci_upper"])

    # --------------------
    # 评估指标（可选）
    # --------------------
    if show_metrics:
        y_true = y.values
        y_pred = df["food_fit"].values

        ss_res = np.sum((y_true - y_pred) ** 2)
        ss_tot = np.sum((y_true - y_true.mean()) ** 2)
        r2_orig = 1 - ss_res / ss_tot

        rmse = np.sqrt(np.mean((y_true - y_pred) ** 2))
        mae = np.mean(np.abs(y_true - y_pred))

        print(model.summary())
        print("-" * 50)
        print(f"Pseudo R² (original scale): {r2_orig:.4f}")
        print(f"Residual variance (log-space): {np.var(model.resid):.4f}")
        print(f"RMSE: {rmse:.4f}")
        print(f"MAE: {mae:.4f}")

    return df, model

import plotly.graph_objects as go

# import plotly.express as ex
def plot_true_pred(df,
                   x_name='Date', 
                   y_true_name='avg_food_kg', true_marker='markers', 
                   y_pred_name='food_fit', pred_marker='lines',
                   plot_title='喂食量时序图'):
    
    fig = go.Figure()
    
    # 实际曲线
    fig.add_trace(
        go.Scatter(
            x=df[x_name],
            y=df[y_true_name],
            mode=true_marker,
            name='喂食量',
            line=dict(color='red', width=2),
            marker=dict(size=4),
            customdata=df[['Date', 'age']],
            hovertemplate=
            '日期: %{customdata[0]:%Y-%m-%d}<br>'
            '日龄: %{customdata[1]} 天<br>'
            '实际喂食量: %{y:.2f} kg'
        )
    )

    # 拟合曲线
    fig.add_trace(
        go.Scatter(
            x=df[x_name],
            y=df[y_pred_name],
            mode=pred_marker,
            name='拟合喂食量',
            line=dict(color='blue', width=2),
            marker=dict(size=4),
            customdata=df[['Date', 'age']],
            hovertemplate=
            '日期: %{customdata[0]:%Y-%m-%d}<br>'
            '日龄: %{customdata[1]} 天<br>'
            '拟合喂食量: %{y:.2f} kg'
        )
    )

    fig.update_layout(
        title=plot_title,
        xaxis_title='日期',
        yaxis_title='喂食量(kg)',
        template='plotly_white',
    )

    # x 轴日期格式 & 网格
    fig.update_xaxes(
        tickformat='%Y-%m-%d',
        tickangle=-90,
        dtick='D1',              # 每天一个刻度
        showgrid=True,
        )
    
    return fig