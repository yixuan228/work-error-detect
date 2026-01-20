import numpy as np

def logistic(x, L, k, x0):
    """
    Logistic 生长模型

    y = L / (1 + exp(-k * (x - x0)))
    """
    return L / (1 + np.exp(-k * (x - x0)))

from scipy.optimize import curve_fit

def fit_logistic(x, y, bounds=None, p0=None):
    """
    多点 Logistic 非线性最小二乘拟合

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
        {'L': L, 'k': k, 'x0': x0, 'cov': pcov}
    """
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    if p0 is None:
        p0 = [
            y.max() * 1.2,   # L
            0.1,             # k
            np.median(x)     # x0
        ]

    if bounds is None:
        bounds = (0, np.inf)

    params, pcov = curve_fit(
        logistic,
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

def fit_logistic_fixed_L(x, y, L):
    """
    固定上限 L 的 Logistic 拟合（线性化解析解）
    """
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    if np.any(y <= 0) or np.any(y >= L):
        raise ValueError("y 必须满足 0 < y < L")

    z = np.log(L / y - 1)

    a, b = np.polyfit(x, z, deg=1)

    k = -a
    x0 = b / k

    return {
        "L": L,
        "k": k,
        "x0": x0
    }

def logistic_pred(x, params):
    """
    Logistic 生长曲线模型

    y = L * exp(-exp(-k * (x - x0)))
    """
    L = params['L']
    k = params['k'] 
    x0 = params['x0'] 

    return L / (1 + np.exp(-k * (x - x0)))