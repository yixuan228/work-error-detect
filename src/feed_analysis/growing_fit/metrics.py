import numpy as np

def get_r2(y_true, y_pred):
    """
    计算 R²（决定系数）
    """
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)

    if y_true.shape != y_pred.shape:
        raise ValueError("y_true 和 y_pred 形状必须一致")

    ss_res = np.sum((y_true - y_pred) ** 2)
    ss_tot = np.sum((y_true - np.mean(y_true)) ** 2)

    return np.nan if ss_tot == 0 else 1 - ss_res / ss_tot

def get_rmse(y_true, y_pred):
    """
    计算 RMSE（均方根误差）
    """
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)

    if y_true.shape != y_pred.shape:
        raise ValueError("y_true 和 y_pred 形状必须一致")
    return np.sqrt(np.mean((y_true - y_pred) ** 2))

def get_mae(y_true, y_pred):
    """
    计算 MAE（平均绝对误差）
    """
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)

    if y_true.shape != y_pred.shape:

        raise ValueError("y_true 和 y_pred 形状必须一致")
    
    return np.mean(np.abs(y_true - y_pred))

import numpy as np
def get_mse(y_true, y_pred):
    """
    计算 MSE（Mean Squared Error）
    """
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)

    if y_true.shape != y_pred.shape:
        raise ValueError("y_true 和 y_pred 形状必须一致")

    return np.mean((y_true - y_pred) ** 2)

def get_mape(y_true, y_pred):
    """
    计算 MAPE（平均绝对百分比误差）
    """
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)

    if y_true.shape != y_pred.shape:
        raise ValueError("y_true 和 y_pred 形状必须一致")
    
    return np.mean(np.abs((y_true - y_pred) / y_true)) * 100

