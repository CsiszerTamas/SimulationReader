package assignment_3;

public class Value {

    private double T;
    private double M;
    private double MM_not_needed;
    private double M2;
    private double chi;
    private double cv;

    public Value(double t, double m, double MM_not_needed, double chi, double cv) {
        this.T = t;
        this.M = m;
        this.MM_not_needed = MM_not_needed;
        this.chi = chi;
        this.cv = cv;
    }

    public double getT() {
        return T;
    }

    public void setT(double t) {
        T = t;
    }

    public double getM() {
        return M;
    }

    public void setM(double m) {
        M = m;
    }

    public double getMM_not_needed() {
        return MM_not_needed;
    }

    public void setMM_not_needed(double MM_not_needed) {
        this.MM_not_needed = MM_not_needed;
    }

    public double getM2() {
        return M2;
    }

    public void setM2(double m2) {
        M2 = m2;
    }

    public double getChi() {
        return chi;
    }

    public void setChi(double chi) {
        this.chi = chi;
    }

    public double getCv() {
        return cv;
    }

    public void setCv(double cv) {
        this.cv = cv;
    }
}
