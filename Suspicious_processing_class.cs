using System;

public class Signal_processing
{
    public static double[] Zero_average_corr(double [] w, int index)
    {
        double[] angle,Av;
        angle = new double[index];
        Av = new double[index];
        angle[0] = 0;
        for (int i = 1; i < index; i++) angle[i] = Simple_Integration(angle[i-1],w[i],0.01);
        for (int i = 0; i < index; i++) Av[i] = Average500(angle,i);
        for (int i = 0; i < index; i++) angle[i] = angle[i] - Av[i];
        return angle;
    }
    private static double Simple_Integration(double A, double B, double dt)
    {        
        return A+B*dt;
    }
    private static double Average50(double[] A, int index)
    {
        double Summ = 0;
        if (index < 25)
        {
            for (int i = 0; i < index + 25; i++)
            {
                Summ += A[i];
            }
            return Summ / (index + 26);
        }
        if (index >= 25 && A.Length - index > 25)
        {
            for (int i = index - 25; i < index + 25; i++)
            {
                Summ += A[i];
            }
            return Summ / (index + 50);
        }
        if (A.Length - index < 25)
        {
            for (int i = index - 25; i < A.Length; i++)
            {
                Summ += A[i];
            }
            return Summ / (A.Length - index + 26);
        }
        return 1;
    }
    private static double Average500(double[] A, int index)
    {
        double Summ = 0;
        if (index < 250)
        {
            for (int i = 0; i < index + 199; i++)
            {
                Summ += A[i];
            }
            return Summ / (index + 198);
        }
        if (index>=250 && A.Length-index > 250)
        {
            for (int i = index - 250; i < index + 250; i++)
            {
                Summ += A[i];
            }
            return Summ / 500;
        }
        if (A.Length - index < 251)
        {
            for (int i = index - 250; i < A.Length; i++)
            {
                Summ += A[i];
            }
            return Summ / (A.Length - index + 249);
        }
        return 1;
    }
}
