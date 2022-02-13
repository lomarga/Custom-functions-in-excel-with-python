from pydoc import doc
from tokenize import Number
import xlwings as xw

from sklearn.linear_model import LinearRegression
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]


@xw.func
def hello(name):
    return f"Hello {name}!"

lr = LinearRegression()

@xw.func
@xw.arg('X_train', np.array, ndim=2, doc="Select X_train")
@xw.arg('y_train', np.array, ndim=2, doc="Select y_train")
@xw.arg('X_test', np.array, ndim=2, doc="Select X_test")
def pred_lr(X_train, y_train, X_test): 
    """
    Creates a linear regression model and predicts the output of the test data
    """
    lr.fit(X_train, y_train)
    y_pred = lr.predict(X_test)
    return y_pred


@xw.func
@xw.arg('X_train', np.array, ndim=2)
@xw.arg('y_train', np.array, ndim=2)
def coefficients_lr(X_train, y_train): 
    """
    Creates a linear regression model and returns the coefficients
    """
    lr.fit(X_train, y_train)
    return lr.coef_

@xw.func
@xw.arg('X_train', np.array, ndim=2)
@xw.arg('y_train', np.array, ndim=2)
def intercept_lr(X_train, y_train): 
    """
    Creates a linear regression model and returns the intercept
    """
    lr.fit(X_train, y_train)
    return lr.intercept_

@xw.func
@xw.arg('X_train', np.array, ndim=2)
@xw.arg('y_train', np.array, ndim=2)
def score_lr(X_train, y_train):
    """
    Creates a linear regression model and returns the R2 score
    """
    lr.fit(X_train, y_train)
    return round(lr.score(X_train, y_train), 3)

@xw.func
@xw.arg('X_train', np.array, ndim=2)
@xw.arg('y_train', np.array, ndim=2)
def rsme_lr(X_train, y_train):
    """
    Creates a linear regression model and returns the RSME
    """
    lr.fit(X_train, y_train)
    return round(np.sqrt(np.mean((lr.predict(X_train) - y_train) ** 2)), 3)


@xw.func
@xw.arg('X_train', pd.DataFrame, index=False, header=False, doc="Select X_train")
@xw.ret(transpose = True)
def kmeans_cluster(X_train):
    """
    Creates a KMeans model and returns the cluster labels
    """
    kmeans = KMeans(n_clusters= int(xw.Range('H1').value), random_state=1984).fit(X_train)
    labels = kmeans.predict(X_train)
    return labels

@xw.func
@xw.arg('X_train', pd.DataFrame, index=False, header=True, doc="Select X_train with column names")
def kmeans_centroids(X_train): 
    """
    Creates a KMeans model and returns the centroids
    """
    kmeans = KMeans(n_clusters= int(xw.Range('H1').value), random_state=1984).fit(X_train)
    centroids = kmeans.cluster_centers_
    centr_kmeans =pd.DataFrame(centroids)
    centr_kmeans.columns = X_train.columns
    centr_kmeans.index.rename('Centroids', inplace=True)
    return centr_kmeans
    


if __name__ == "__main__":
    xw.Book("PythonFunctions.xlsm").set_mock_caller()
    main()
