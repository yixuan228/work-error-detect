import numpy as np
from scipy.optimize import fsolve

def gompertz(x, L, k, x0):
    """
    Gompertz 生长曲线模型

    y = L * exp(-exp(-k * (x - x0)))
    """

    return L * np.exp(-np.exp(-k * (x - x0)))

# 根据三点严格求解Gompertz参数
def gompertz_equations(params, x, y):
    """
    Gompertz 三参数方程组（用于 3 点严格求解）

    """

    L, k, x0 = params
    return [
        gompertz(x[0], L, k, x0) - y[0],
        gompertz(x[1], L, k, x0) - y[1],
        gompertz(x[2], L, k, x0) - y[2],
    ]

def fit_gompertz_through_3points(x, y, init=None):
    """
    用 3 个点严格求解 Gompertz 参数（曲线必过 3 点）

    Returns
    -------
    params : dict
        {'L': L, 'k': k, 'x0': x0}
    """
    
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    if len(x) != 3 or len(y) != 3:
        raise ValueError("必须恰好提供 3 个点")

    # 初值非常重要
    if init is None:
        L0 = y.max() * 1.2
        k0 = 0.1
        x00 = np.mean(x)
        init = [L0, k0, x00]

    sol, info, ier, msg = fsolve(
        gompertz_equations,
        init,
        args=(x, y),
        full_output=True
    )

    if ier != 1:
        raise RuntimeError(f"Gompertz 求解失败: {msg}")

    L, k, x0 = sol
    return {"L": L, "k": k, "x0": x0}


# 定x0两点求解 Gompertz 参数
import numpy as np
from scipy.optimize import curve_fit
def fit_gompertz_fixed_x0(x, y, x0):
    """
    固定拐点 x0，拟合 Gompertz 的 L 和 k（最小二乘）
        
    Returns
    -------
    params : dict
        {'L': L, 'k': k, 'x0': x0}
    """
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    if np.any(y <= 0):
        raise ValueError("y 必须 > 0")

    # 固定 x0 的包装函数
    def gompertz_fixed_x0(x, L, k):
        return gompertz(x, L, k, x0)

    # 初值（非常重要）
    L0 = y.max() * 1.2
    k0 = 0.05

    params, _ = curve_fit(
        gompertz_fixed_x0,
        x,
        y,
        p0=[L0, k0],
        maxfev=10000
    )

    L, k = params
    return {"L": L, "k": k, "x0": x0}

import numpy as np

def fit_gompertz_fixed_L(x, y, L):
    """
    固定上限 L，线性化求解 Gompertz 的 k 和 x0

    Returns
    -------
    params : dict
        {'L': L, 'k': k, 'x0': x0}
    """
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    if np.any(y <= 0) or np.any(y >= L):
        raise ValueError("y 必须满足 0 < y < L")

    # 线性化变换
    z = np.log(-np.log(y / L))

    # 线性回归：z = a * x + b
    a, b = np.polyfit(x, z, deg=1)

    k = -a
    x0 = b / k

    return {
        "L": L,
        "k": k,
        "x0": x0
    }

import numpy as np
from scipy.optimize import curve_fit
def fit_gompertz(x, y, bounds=None, p0=None):
    """
    多点 Gompertz 非线性最小二乘拟合
    
    Parameters
    ----------
    x, y : array-like
    bounds : tuple, optional
        参数上下界 ((L_min, k_min, x0_min), (L_max, k_max, x0_max))
    p0 : list, optional
        初始值 [L, k, x0]
        
    Returns
    -------
    params : dict
        {'L': L, 'k': k, 'x0': x0}
    """

    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    if p0 is None:
        p0 = [
            y.max() * 1.2,   # L
            0.05,            # k
            np.median(x)     # x0
        ]

    if bounds is None:
        bounds = (0, np.inf)

    params, pcov = curve_fit(
        gompertz,
        x,
        y,
        p0=p0,
        bounds=bounds,
        maxfev=20000
    )

    L, k, x0 = params
    return {
        "L": L,
        "k": k,
        "x0": x0,
        "cov": pcov
    }

def gompertz_pred(x, params):
    """
    Gompertz 生长曲线模型

    y = L * exp(-exp(-k * (x - x0)))
    """
    L = params['L']
    k = params['k'] 
    x0 = params['x0'] 

    return L * np.exp(-np.exp(-k * (x - x0)))