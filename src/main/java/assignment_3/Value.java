package assignment_3;

public class Value {

    private double T;
    private double M;
    private double E;
    private double M2;
    private double chi;
    private double cv;

    public Value(double t, double m, double E, double chi, double cv) {
        this.T = t;
        this.M = m;
        this.E = E;
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

    public double getE() {
        return E;
    }

    public void setE(double e) {
        this.E = e;
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
