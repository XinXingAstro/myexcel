package com.xinxing.excel;

public class ExcelDataPO {
    private String projectNumber;
    private int devDeploy;
    private int testDeploy;

    public String getProjectNumber() {
        return projectNumber;
    }

    public void setProjectNumber(String projectNumber) {
        this.projectNumber = projectNumber;
    }

    public int getDevDeploy() {
        return devDeploy;
    }

    public void setDevDeploy(int devDeploy) {
        this.devDeploy = devDeploy;
    }

    public int getTestDeploy() {
        return testDeploy;
    }

    public void setTestDeploy(int testDeploy) {
        this.testDeploy = testDeploy;
    }
}
