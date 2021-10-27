# SKU Quality Control Application

## Introduction

The purpose of SKU Quality Control application is to investigate the atypical purchases that are
recorded in one audit period.
The process of finding atypical purchases reaches the SKU level per store
of the _Emrc Retail Audit sample._

The distance measure which is used to implement the above, based on Shanon’s theory, is the
following:

<img src="https://latex.codecogs.com/gif.latex?%5Ctext%7B%20%7DD%28P_%7Bt&plus;1%7D%2CP_t%29%3DP_%7Bt&plus;1%7D%5Ccdot%5Cln%28P_%7Bt&plus;1%7D/P_t%29&plus;P_t-P_%7Bt&plus;1%7D%5Ctext%7B%20%7D" />

## An example

A typical example that describes this property is the following:
We have 2 stores where the purchases for the same SKU in 2 consecutive periods are

<img src="https://latex.codecogs.com/gif.latex?%5Ctext%7Bstore%20A%3A%20%7DP_t%3D1%2C%20P_%7Bt&plus;1%7D%3D2%5Ctext%7B%2C%20and%20store%20B%3A%20%7DP_t%3D5%2C%20P_%7Bt&plus;1%7D%3D10." />

We immediately observe the doubling of the purchases in the period t + 1, that is, we have a 100%
increase in both stores. Here, the increase in store A should not worry us because it is reasonable and
expected to buy an additional SKU and we must pay attention to store B, something that the classic
distance measures do not indicate.
However, by using the above type of distance, we obtain

<img src="https://latex.codecogs.com/gif.latex?D%28P_%7Bt&plus;1%7D%2CP_t%29%3D0.386%5Ctext%7B%20for%20the%20store%20A%20and%20%7DD%28P_%7Bt&plus;1%7D%2CP_t%29%3D1.931%5Ctext%7B%20for%20the%20store%20B%2C%7D" />

## Gui
