import xlwings as xw

from sklearn.linear_model import LinearRegression
import numpy as np

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]


@xw.func
def hello(name):
    return f"Hello {name}!"

@xw.func
@xw.arg('X_train', np.array, ndim=2)
@xw.arg('y_train', np.array, ndim=2)
@xw.arg('X_test', np.array, ndim=2)
def pred_lr(X_train, y_train, X_test): #Create a linear regression model
    lr = LinearRegression()
    lr.fit(X_train, y_train)
    y_pred = lr.predict(X_test)
    return y_pred


@xw.func
@xw.arg('X_train', np.array, ndim=2)
@xw.arg('y_train', np.array, ndim=2)
def coefficients_lr(X_train, y_train): #Returns the coefficients
    lr = LinearRegression()
    lr.fit(X_train, y_train)
    return lr.coef_

@xw.func
@xw.arg('X_train', np.array, ndim=2)
@xw.arg('y_train', np.array, ndim=2)
def intercept_lr(X_train, y_train): #Returns the intercept
    lr = LinearRegression()
    lr.fit(X_train, y_train)
    return lr.intercept_

@xw.func
@xw.arg('X_train', np.array, ndim=2)
@xw.arg('y_train', np.array, ndim=2)
def score_lr(X_train, y_train): #Returns the score
    lr = LinearRegression()
    lr.fit(X_train, y_train)
    return round(lr.score(X_train, y_train), 3)

@xw.func
@xw.arg('X_train', np.array, ndim=2)
@xw.arg('y_train', np.array, ndim=2)
def rsme_lr(X_train, y_train): #Returns the root mean squared error
    lr = LinearRegression()
    lr.fit(X_train, y_train)
    return round(np.sqrt(np.mean((lr.predict(X_train) - y_train) ** 2)), 3)





if __name__ == "__main__":
    xw.Book("linearRegression.xlsm").set_mock_caller()
    main()
