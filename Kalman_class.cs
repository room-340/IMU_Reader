using System;
using MathNet.Numerics.LinearAlgebra.Double;

public static class Kalman_class
{
    public const double ACCLERATION_NOISE = 1;//20000;
    public const double MAGNETIC_FIELD_NOISE = 1;
    public const double ANGULAR_VELOCITY_NOISE = 0.001;
    public struct Parameters
    {
        public double dT;
        public double declination;
        public double g;
        public double accl_threshold;
        public double field;
        public double field_threshold;
        public DenseVector accl_coefs;
        public DenseVector magn_coefs;
        public DenseVector gyro_coefs;
        public double accl_noise;
        public double magn_noise;
        public double scale_noise;
        public double skew_noise;


        public Parameters(DenseVector Accl_coefs, DenseVector Magn_coefs, DenseVector Gyro_coefs)
        {
            dT = (double)1/100;
            declination = (double)10 * Math.PI / 180;
            g = (double) 1;
            accl_threshold = (double) 0.1;
            field = (double) 1;
            field_threshold = (double) 0.1;
            accl_coefs = Accl_coefs;
            magn_coefs = Magn_coefs;
            gyro_coefs = Gyro_coefs;
            accl_noise = ACCLERATION_NOISE;
            magn_noise = MAGNETIC_FIELD_NOISE;
            scale_noise = (double) Math.Pow(10,-15);
            skew_noise = (double) Math.Pow(10,-15);
        }
    }

    public struct Sensors
    {
        public Matrix w;
        public Matrix a;
        public Matrix m;

        public Sensors(Matrix W, Matrix A, Matrix M)
        {
            w = W;
            a = A;
            m = M;
        }
    }

    public struct State
    {
        public Matrix R;
        public Matrix Q;
        public Matrix P;
        public Matrix dw;
        public Matrix dB;
        public Matrix dG;
        public Matrix q;

        public State(double accl_noise, double magn_noise, double angle_noise, double bias_noise, double scale_noise, double skew_noise, Matrix Quat)
        {
            System.Func<int, int, double> matrix_fill = (x, y) => 0;
            R = new DiagonalMatrix(3, 3, accl_noise);
            
            R.At(2, 2, magn_noise);
            
            Q = DenseMatrix.Create(15, 15, matrix_fill);                   
            
            Q.At(0, 0, angle_noise);
            Q.At(1, 1, angle_noise);
            Q.At(2, 2, angle_noise);

            Q.At(3, 3, bias_noise);
            Q.At(4, 4, bias_noise);
            Q.At(5, 5, bias_noise);

            Q.At(6, 6, scale_noise);
            Q.At(7, 7, scale_noise);
            Q.At(8, 8, scale_noise);

            Q.At(9, 9, skew_noise);
            Q.At(10, 10, skew_noise);
            Q.At(11, 11, skew_noise);

            Q.At(12, 12, skew_noise);
            Q.At(13, 13, skew_noise);
            Q.At(14, 14, skew_noise);

            P = DenseMatrix.Create(15, 15, matrix_fill);            

            P.At(0, 0, (double)Math.Pow(10, -1));
            P.At(1, 1, (double)Math.Pow(10, -1));
            P.At(2, 2, (double)Math.Pow(10, -1));

            P.At(3, 3, (double)Math.Pow(10, -3));
            P.At(4, 4, (double)Math.Pow(10, -3));
            P.At(5, 5, (double)Math.Pow(10, -3));

            P.At(6, 6, (double)Math.Pow(10, -8));
            P.At(7, 7, (double)Math.Pow(10, -8));
            P.At(8, 8, (double)Math.Pow(10, -8));

            P.At(9, 9, (double)Math.Pow(10, -8));
            P.At(10, 10, (double)Math.Pow(10, -8));
            P.At(11, 11, (double)Math.Pow(10, -8));

            P.At(12, 12, (double)Math.Pow(10, -8));
            P.At(13, 13, (double)Math.Pow(10, -8));
            P.At(14, 14, (double)Math.Pow(10, -8));

            dw = DenseMatrix.Create(3, 1, matrix_fill);

            dB = DenseMatrix.Create(3, 1, matrix_fill);

            dG = DenseMatrix.Create(6, 1, matrix_fill);

            q = Quat;          
        }
    }


    public static Tuple <Vector, Sensors, State> AHRS_LKF_EULER(Sensors Sense, State State, Parameters Param)
    {
        System.Func<int, double> vect_fill = (x) => 0;
        System.Func<int, int, double> matrix_fill = (x, y) => 0;

        Vector Attitude = DenseVector.Create(6, vect_fill);
        bool restart = false;
        // get sensor data
        Matrix m = Sense.m;
        Matrix a = Sense.a;
        Matrix w = Sense.w;    
    
        // Correct magntometers using callibration coefficients
        Matrix B = new DenseMatrix(3,3);
        B.At(0, 0, Param.magn_coefs.At(0));
        B.At(0, 1, Param.magn_coefs.At(3));
        B.At(0, 2, Param.magn_coefs.At(4));
        B.At(1, 0, Param.magn_coefs.At(5));
        B.At(1, 1, Param.magn_coefs.At(1));
        B.At(1, 2, Param.magn_coefs.At(6));
        B.At(2, 0, Param.magn_coefs.At(7));
        B.At(2, 1, Param.magn_coefs.At(8));
        B.At(2, 2, Param.magn_coefs.At(2));
        Matrix B0 = new DenseMatrix(3, 1);
        B0.At(0, 0, Param.magn_coefs.At(9));
        B0.At(1, 0, Param.magn_coefs.At(10));
        B0.At(2, 0, Param.magn_coefs.At(11));        

        m = Matrix_Transpose(Matrix_Mult(Matrix_Minus(new DiagonalMatrix(3,3,1),B),Matrix_Minus(Matrix_Transpose(m),B0)));

        // Correct accelerometers using callibration coefficients
        B.At(0, 0, Param.accl_coefs.At(0));
        B.At(0, 1, Param.accl_coefs.At(3));
        B.At(0, 2, Param.accl_coefs.At(4));
        B.At(1, 0, Param.accl_coefs.At(5));
        B.At(1, 1, Param.accl_coefs.At(1));
        B.At(1, 2, Param.accl_coefs.At(6));
        B.At(2, 0, Param.accl_coefs.At(7));
        B.At(2, 1, Param.accl_coefs.At(8));
        B.At(2, 2, Param.accl_coefs.At(2));

        B0.At(0, 0, Param.accl_coefs.At(9));
        B0.At(1, 0, Param.accl_coefs.At(10));
        B0.At(2, 0, Param.accl_coefs.At(11));

        a = Matrix_Transpose(Matrix_Mult(Matrix_Minus(new DiagonalMatrix(3, 3, 1), B), Matrix_Minus(Matrix_Transpose(a), B0)));

        // Correct gyroscopes using callibration coefficients
        B.At(0, 0, Param.gyro_coefs.At(0));
        B.At(0, 1, Param.gyro_coefs.At(3));
        B.At(0, 2, Param.gyro_coefs.At(4));
        B.At(1, 0, Param.gyro_coefs.At(5));
        B.At(1, 1, Param.gyro_coefs.At(1));
        B.At(1, 2, Param.gyro_coefs.At(6));
        B.At(2, 0, Param.gyro_coefs.At(7));
        B.At(2, 1, Param.gyro_coefs.At(8));
        B.At(2, 2, Param.gyro_coefs.At(2));

        B0.At(0, 0, Param.gyro_coefs.At(9));
        B0.At(1, 0, Param.gyro_coefs.At(10));
        B0.At(2, 0, Param.gyro_coefs.At(11));

        w = Matrix_Transpose(Matrix_Mult(Matrix_Minus(new DiagonalMatrix(3, 3, 1), B), Matrix_Minus(Matrix_Transpose(w), B0)));

        // Get State
        Matrix q = State.q;
        Matrix dB = State.dB;
        Matrix dG = State.dG;
        Matrix dw = State.dw;
        Matrix P = State.P;

        Matrix Wb = Matrix_Transpose(w);
        Matrix Ab = Matrix_Transpose(a);
        Matrix Mb = Matrix_Transpose(m);

        double dT = Param.dT;
        
        //Correct Gyroscopes for estimate biases and scale factor
        B.At(0, 0, dB.At(0, 0));
        B.At(0, 1, dG.At(0, 0));
        B.At(0, 2, dG.At(1, 0));
        B.At(1, 0, dG.At(2, 0));
        B.At(1, 1, dB.At(1, 0));
        B.At(1, 2, dG.At(3, 0));
        B.At(2, 0, dG.At(4, 0));
        B.At(2, 1, dG.At(5, 0));
        B.At(2, 2, dB.At(2, 0));


        Matrix Omegab_ib;
        Omegab_ib = Matrix_Minus(Matrix_Mult(Matrix_Minus(new DiagonalMatrix(3, 3, 1), B), Wb),dw);

        if (q.At(0, 0).ToString() == "NaN")
        {
            restart = true;
        }

        //Quternion calculation
        q = mrotate(q, Omegab_ib, dT);

        if (q.At(0,0).ToString() == "NaN")
        {
            restart = true;
        }

        //DCM calculation
        Matrix Cbn = quat_to_DCM(q);

        //Gyro Angles
        Matrix angles = dcm2angle(Cbn);
        double Psi = (double) angles.At(0, 0);
        double Theta = (double) angles.At(0, 1);
        double Gamma = (double) angles.At(0, 2);        

        //Acceleration Angles
        double ThetaAcc = (double)Math.Atan2(Ab.At(0, 0), Math.Sqrt(Ab.At(1, 0) * Ab.At(1, 0) + Ab.At(2, 0) * Ab.At(2, 0)));
        double GammaAcc = (double)-Math.Atan2(Ab.At(1, 0), Math.Sqrt(Ab.At(0, 0) * Ab.At(0, 0) + Ab.At(2, 0) * Ab.At(2, 0)));
        
        //Horizontal projection of magnetic field  
        angles.At(0, 0, 0);
        angles.At(0, 1, ThetaAcc);
        angles.At(0, 2, GammaAcc);
        Matrix Cbh = angle2dcm(angles);
        Matrix Mh = Matrix_Mult(Matrix_Transpose(Cbh), Mb);

        //Magnetic Heading
        double PsiMgn = (double)(-Math.Atan2(Mh.At(1,0),Mh.At(0,0)) + Param.declination);

        //System matrix
        Matrix A = DenseMatrix.Create(15, 15, matrix_fill);
        Matrix I = new DiagonalMatrix (15, 15, 1);

        A.At(0, 3, 1);
        A.At(1, 4, 1);
        A.At(2, 5, 1);

        A.At(0, 6, 1);
        A.At(1, 7, 1);
        A.At(2, 8, 1);

        A.At(0, 9, Omegab_ib.At(1, 0));
        A.At(0, 10, Omegab_ib.At(2, 0));
        A.At(1, 11, Omegab_ib.At(0, 0));                
        A.At(1, 12, Omegab_ib.At(2, 0));
        A.At(2, 13, Omegab_ib.At(0, 0));
        A.At(2, 14, Omegab_ib.At(1, 0));

        Matrix F = Matrix_Plus(I, Matrix_Const_Mult(A,dT));

        //Measurment Matrix
        double dPsi = Psi - PsiMgn;
        if (dPsi > Math.PI) dPsi = (double)(dPsi - 2 * Math.PI);
        if (dPsi < -Math.PI) dPsi = (double)(dPsi + 2 * Math.PI);

        double dTheta = Theta - ThetaAcc;
        if (dTheta > Math.PI) dTheta = (double)(dTheta - 2 * Math.PI);
        if (dTheta < -Math.PI) dTheta = (double)(dTheta + 2 * Math.PI);

        double dGamma = Gamma - GammaAcc;
        if (dGamma > Math.PI) dGamma = (double)(dGamma - 2 * Math.PI);
        if (dGamma < -Math.PI) dGamma = (double)(dGamma + 2 * Math.PI);

        Matrix z = DenseMatrix.Create(3,1,matrix_fill);
        z[0, 0] = 1; z[1, 0] = 1; z[2, 0] = 1;

        z.At(0, 0, dGamma);
        z.At(1, 0, dTheta);
        z.At(2, 0, dPsi);

        Matrix H = DenseMatrix.Create(3,15,matrix_fill);
        H.At(0, 0, 1);
        H.At(1, 1, 1);
        H.At(2, 2, 1); // MAGNETOMETER CORRECTION

        if (Math.Abs(Math.Sqrt(Math.Pow(Ab.At(0, 0), 2) + Math.Pow(Ab.At(1, 0), 2) + Math.Pow(Ab.At(2, 0), 2))) - 1 > Param.accl_threshold)
        {
            H.At(0, 0, 0);
            H.At(1, 1, 0);
        }

        if (Math.Abs(Math.Sqrt(Math.Pow(Mh.At(0, 0), 2) + Math.Pow(Mh.At(1, 0), 2) + Math.Pow(Mh.At(2, 0), 2))) - 1 > Param.field_threshold)
        {
            H.At(2, 2, 0);
        }

        //Kalman Filter
        Matrix Q = State.Q;
        Matrix R = State.R;

        P = Matrix_Plus(Matrix_Mult(Matrix_Mult(F, P), Matrix_Transpose(F)),Q);
        
        if (P.At(0, 0).ToString() == "NaN")
        {
            P = new DiagonalMatrix(15, 15, (double)Math.Pow(10, -8));

            P.At(0, 0, (double)Math.Pow(10, -1));
            P.At(1, 1, (double)Math.Pow(10, -1));
            P.At(2, 2, (double)Math.Pow(10, -1));

            P.At(3, 3, (double)Math.Pow(10, -3));
            P.At(4, 4, (double)Math.Pow(10, -3));
            P.At(5, 5, (double)Math.Pow(10, -3));
            restart = true;
        }

        Tuple<Matrix, Matrix> KF_result;
        KF_result = KF_Cholesky_update(P,z,R,H);
        Matrix xf = KF_result.Item1;
        P = KF_result.Item2;

        Matrix df_hat = DenseMatrix.Create(3,1,matrix_fill);;
        df_hat.At(0, 0, xf.At(0, 0));
        df_hat.At(1, 0, xf.At(1, 0));
        df_hat.At(2, 0, xf.At(2, 0));
        

        Matrix dw_hat = DenseMatrix.Create(3,1,matrix_fill);
        dw_hat.At(0, 0, xf.At(3, 0));
        dw_hat.At(1, 0, xf.At(4, 0));
        dw_hat.At(2, 0, xf.At(5, 0));
        Matrix dB_hat = DenseMatrix.Create(3,1,matrix_fill);
        dB_hat.At(0, 0, xf.At(6, 0));
        dB_hat.At(1, 0, xf.At(7, 0));
        dB_hat.At(2, 0, xf.At(8, 0));
        Matrix dG_hat = DenseMatrix.Create(6,1,matrix_fill);
        dG_hat.At(0, 0, xf.At(9, 0));
        dG_hat.At(1, 0, xf.At(10, 0));
        dG_hat.At(2, 0, xf.At(11, 0));
        dG_hat.At(3, 0, xf.At(12, 0));
        dG_hat.At(4, 0, xf.At(13, 0));
        dG_hat.At(5, 0, xf.At(14, 0));

        dw = Matrix_Plus(dw, dw_hat);
        dB = Matrix_Plus(dB, dB_hat);
        dG = Matrix_Plus(dG, dG_hat);



        Matrix dCbn = DenseMatrix.Create(3,3,matrix_fill);

        dCbn.At(0, 1, -df_hat.At(2, 0));
        dCbn.At(0, 2, df_hat.At(1, 0));
        dCbn.At(1, 0, df_hat.At(2, 0));
        dCbn.At(1, 2, -df_hat.At(0, 0));
        dCbn.At(2, 0, -df_hat.At(1, 0));
        dCbn.At(2, 1, df_hat.At(0, 0));

        Cbn = Matrix_Mult(Matrix_Plus(new DiagonalMatrix(3, 3, 1), dCbn), Cbn);
        if (Cbn.At(0, 0).ToString() == "NaN")
        {
            restart = true;
        }
        if (dcm2quat(Cbn).At(0, 0).ToString() == "NaN")
        {
            restart = true;
        }
        q = dcm2quat(Cbn);
        if (q.At(0, 0).ToString() == "NaN")
        {
            restart = true;
        }
        q = quat_norm(q);

        if (q.At(0, 0).ToString() == "NaN")
        {
            restart = true;
        }
        
        Attitude.At(0, Psi);
        Attitude.At(1, Theta);
        Attitude.At(2, Gamma);
        Attitude.At(3, PsiMgn);
        Attitude.At(4, ThetaAcc);
        Attitude.At(5, GammaAcc);

        State.q = q;
        State.dG = dG;
        State.dB = dB;
        State.dw = dw;
        State.P = P;


        Sense.w = Matrix_Transpose(Omegab_ib);
        Sense.a = a;
        Sense.m = m;

        if (restart)
        {
            Matrix Initia_quat = DenseMatrix.Create(1, 4, matrix_fill);
            Initia_quat.At(0, 0, 1);
            State = new Kalman_class.State(ACCLERATION_NOISE, MAGNETIC_FIELD_NOISE, ANGULAR_VELOCITY_NOISE,
                Math.Pow(10, -6), Math.Pow(10, -15), Math.Pow(10, -15), Initia_quat);
        }

        return new Tuple <Vector, Sensors, State>(Attitude, Sense, State); 
    }

    public static Matrix Matrix_Minus(Matrix A, Matrix B)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        Matrix C = DenseMatrix.Create(A.RowCount,A.ColumnCount, matrix_fill);
        int i, j;
        if (A.RowCount != B.RowCount | A.ColumnCount != B.ColumnCount) return C;
        for (i = 0; i < A.RowCount; i++)
        {
            for (j = 0; j < A.ColumnCount; j++)
            {
                C.At(i,j,A.At(i,j)-B.At(i,j));
            }
        }
            return C;
    }

    public static Matrix Matrix_Plus(Matrix A, Matrix B)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        Matrix C = DenseMatrix.Create(A.RowCount, A.ColumnCount, matrix_fill);
        int i, j;
        if (A.RowCount != B.RowCount | A.ColumnCount != B.ColumnCount) return C;
        for (i = 0; i < A.RowCount; i++)
        {
            for (j = 0; j < A.ColumnCount; j++)
            {
                C.At(i, j, A.At(i, j) + B.At(i, j));
            }
        }
        return C;
    }

    public static Matrix Matrix_Const_Mult(Matrix A, double B)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        Matrix C = DenseMatrix.Create(A.RowCount, A.ColumnCount, matrix_fill);
        int i, j;
        for (i = 0; i < A.RowCount; i++)
        {
            for (j = 0; j < A.ColumnCount; j++)
            {
                C.At(i, j, A.At(i, j)*B);
            }
        }
        return C;
    }

    public static Matrix Matrix_Mult(Matrix A, Matrix B)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        Matrix C = DenseMatrix.Create(A.RowCount, B.ColumnCount, matrix_fill);
        int i, j, k;
        double el;
        if (A.ColumnCount != B.RowCount) return C;
        for (i = 0; i < A.RowCount; i++)
        {
            for (j = 0; j < B.ColumnCount; j++)
            {
                el = 0;
                for (k = 0; k < B.RowCount; k++)
                {
                    el = el + A.At(i, k) * B.At(k, j);
                }
                C.At(i, j, el);
            }
        }
        return C;
    }

    public static Matrix Matrix_Transpose(Matrix A)
    {
        Matrix C = new DenseMatrix(A.ColumnCount, A.RowCount);
        int i, j;
        for (i = 0; i < A.RowCount; i++)
        {
            for (j = 0; j < A.ColumnCount; j++)
            {
                C.At(j, i, A.At(i, j));
            }
        }
        return C;
    }

    public static Matrix mrotate(Matrix q, Matrix w, double dT)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        double Fx = w.At(0, 0) * dT;
        double Fy = w.At(1, 0) * dT;
        double Fz = w.At(2, 0) * dT;
        double Fm = (double)Math.Sqrt(Fx * Fx + Fy * Fy + Fz * Fz);
        double sinFm2 = (double)Math.Sin(Fm / 2);
        double cosFm2 = (double)Math.Cos(Fm / 2);
        Matrix Nd = DenseMatrix.Create(1, 4, matrix_fill);
        if (Fm != (double)0)
        {
            Nd.At(0, 0, cosFm2);
            Nd.At(0, 1, (Fx / Fm) * sinFm2);
            Nd.At(0, 2, (Fy / Fm) * sinFm2);
            Nd.At(0, 3, (Fz / Fm) * sinFm2);
        }
        else Nd.At(0, 0, 1);
        q = quat_multi(q, Nd);
        return q;
    }

    public static Matrix quat_multi(Matrix A, Matrix B)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        Matrix C = DenseMatrix.Create(A.RowCount, A.ColumnCount, matrix_fill);
        if (A.RowCount != 1 | A.ColumnCount != 4) return C;
        if (B.RowCount != 1 | B.ColumnCount != 4) return C;
        C.At(0, 0, A.At(0, 0) * B.At(0, 0) - B.At(0, 1) * A.At(0, 1) - B.At(0, 2) * A.At(0, 2) - B.At(0, 3) * A.At(0, 3));
        C.At(0, 1, A.At(0, 0) * B.At(0, 1) + B.At(0, 0) * A.At(0, 1) + B.At(0, 3) * A.At(0, 2) - B.At(0, 2) * A.At(0, 3));
        C.At(0, 2, A.At(0, 0) * B.At(0, 2) + B.At(0, 0) * A.At(0, 2) - B.At(0, 3) * A.At(0, 1) + B.At(0, 1) * A.At(0, 3));
        C.At(0, 3, A.At(0, 0) * B.At(0, 3) + B.At(0, 2) * A.At(0, 1) - B.At(0, 1) * A.At(0, 2) + B.At(0, 0) * A.At(0, 3));
        return C;
    }

    public static Matrix quat_to_DCM(Matrix q)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        Matrix DCM = DenseMatrix.Create(3, 3, matrix_fill);
        q = quat_norm(q);
        DCM.At(0, 0, (q.At(0, 0) * q.At(0, 0) + q.At(0, 1) * q.At(0, 1) - q.At(0, 2) * q.At(0, 2) - q.At(0, 3) * q.At(0, 3)));
        DCM.At(0, 1, 2 * (q.At(0, 1) * q.At(0, 2) + q.At(0, 0) * q.At(0, 3)));
        DCM.At(0, 2, 2 * (q.At(0, 1) * q.At(0, 3) - q.At(0, 0) * q.At(0, 2)));
        DCM.At(1, 0, 2 * (q.At(0, 1) * q.At(0, 2) - q.At(0, 0) * q.At(0, 3)));
        DCM.At(1, 1, (q.At(0, 0) * q.At(0, 0) - q.At(0, 1) * q.At(0, 1) + q.At(0, 2) * q.At(0, 2) - q.At(0, 3) * q.At(0, 3)));
        DCM.At(1, 2, 2 * (q.At(0, 2) * q.At(0, 3) + q.At(0, 0) * q.At(0, 1)));
        DCM.At(2, 0, 2 * (q.At(0, 1) * q.At(0, 3) + q.At(0, 0) * q.At(0, 2)));
        DCM.At(2, 1, 2 * (q.At(0, 2) * q.At(0, 3) - q.At(0, 0) * q.At(0, 1)));
        DCM.At(2, 2, (q.At(0, 0) * q.At(0, 0) - q.At(0, 1) * q.At(0, 1) - q.At(0, 2) * q.At(0, 2) + q.At(0, 3) * q.At(0, 3)));
        return DCM;
    }

    public static Matrix quat_norm(Matrix q)
    {
        double Mod = (double) Math.Sqrt(q.At(0, 0) * q.At(0, 0) + q.At(0, 1) * q.At(0, 1) + q.At(0, 2) * q.At(0, 2) + q.At(0, 3) * q.At(0, 3));        
        q.At(0, 0, q.At(0, 0) / Mod);
        q.At(0, 1, q.At(0, 1) / Mod);
        q.At(0, 2, q.At(0, 2) / Mod);
        q.At(0, 3, q.At(0, 3) / Mod);
        return q;
    }

    public static Matrix dcm2angle(Matrix DCM)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        Matrix angles = DenseMatrix.Create(1,3,matrix_fill);
        if (DCM.RowCount != 3 | DCM.ColumnCount != 3) return angles;
        angles.At(0, 0, (double)(Math.Atan2(DCM.At(0, 1), DCM.At(0, 0))));
        angles.At(0, 1, (double)-(Math.Atan2(DCM.At(0, 2), Math.Sqrt(1 - DCM.At(0, 2) * DCM.At(0, 2)))));
        angles.At(0, 2, (double)(Math.Atan2(DCM.At(1, 2), DCM.At(2, 2))));
        return angles;
    }

    public static Matrix angle2dcm(Matrix angles)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        Matrix DCM = DenseMatrix.Create(3, 3, matrix_fill);
        if (angles.RowCount != 1 | angles.ColumnCount != 3) return DCM;
        double Cos1 = (double)Math.Cos(angles.At(0, 0));
        double Cos2 = (double)Math.Cos(angles.At(0, 1));
        double Cos3 = (double)Math.Cos(angles.At(0, 2));
        double Sin1 = (double)Math.Sin(angles.At(0, 0));
        double Sin2 = (double)Math.Sin(angles.At(0, 1));
        double Sin3 = (double)Math.Sin(angles.At(0, 2));
        DCM.At(0, 0, Cos2 * Cos1);
        DCM.At(0, 1, Cos2 * Sin1);
        DCM.At(0, 2, -Sin2);
        DCM.At(1, 0, Sin3 * Sin2 * Cos1 - Cos3 * Sin1);
        DCM.At(1, 1, Sin3 * Sin2 * Sin1 + Cos3 * Cos1);
        DCM.At(1, 2, Sin3 * Cos2);
        DCM.At(2, 0, Cos3 * Sin2 * Cos1 + Sin3 * Sin1);
        DCM.At(2, 1, Cos3 * Sin2 * Sin1 - Sin3 * Cos1);
        DCM.At(2, 2, Cos3 * Cos2);
        return DCM;
    }

    public static Matrix dcm2quat(Matrix DCM)
    {
        System.Func<int, int, double> matrix_fill = (x, y) => 0;
        Matrix q = DenseMatrix.Create(1, 4, matrix_fill);
        double tr = trace (DCM);
        if (tr > 0)
        {
            double sqtrp1 = (double)Math.Sqrt(tr + 1);
            q.At(0, 0, (double)0.5 * sqtrp1);
            q.At(0, 1, (DCM.At(1, 2) - DCM.At(2, 1)) / (2 * sqtrp1));
            q.At(0, 2, (DCM.At(2, 0) - DCM.At(0, 2)) / (2 * sqtrp1));
            q.At(0, 3, (DCM.At(0, 1) - DCM.At(1, 0)) / (2 * sqtrp1));
            return q;
        }
        else
        {
            Matrix d = DenseMatrix.Create(1,3,matrix_fill);

            d.At(0, 0, DCM.At(0, 0));
            d.At(0, 1, DCM.At(1, 1));
            d.At(0, 2, DCM.At(2, 2));

            if ((d.At(0,1)>d.At(0,0)&(d.At(0,1)>d.At(0,2))))
            {
                double sqdip1 = (double)Math.Sqrt(d.At(0, 1) - d.At(0, 0) - d.At(0, 2) + 1);
                q.At(0, 2, 0.5 * sqdip1);

                if (sqdip1 != 0)
                {
                    sqdip1 = 0.5 / sqdip1;
                }

                q.At(0, 0, (DCM.At(2, 0) - DCM.At(0, 2)) * sqdip1);
                q.At(0, 1, (DCM.At(0, 1) + DCM.At(1, 0)) * sqdip1);
                q.At(0, 3, (DCM.At(1, 2) + DCM.At(2, 1)) * sqdip1);
            }
            else if (d.At(0, 2) > d.At(0, 0))
            {
                double sqdip1 = (double)Math.Sqrt(d.At(0, 2) - d.At(0, 0) - d.At(0, 1) + 1);
                q.At(0, 3, 0.5 * sqdip1);

                if (sqdip1 != 0)
                {
                    sqdip1 = 0.5 / sqdip1;
                }

                q.At(0, 0, (DCM.At(0, 1) - DCM.At(1, 0)) * sqdip1);
                q.At(0, 1, (DCM.At(2, 0) + DCM.At(0, 2)) * sqdip1);
                q.At(0, 2, (DCM.At(1, 2) + DCM.At(2, 1)) * sqdip1);
            }
            else
            {
                double sqdip1 = (double)Math.Sqrt(d.At(0, 0) - d.At(0, 1) - d.At(0, 2) + 1);
                q.At(0, 1, 0.5 * sqdip1);
                
                if (sqdip1 != 0)
                {
                    sqdip1 = 0.5 * sqdip1;
                }

                q.At(0, 0, (DCM.At(1, 2) - DCM.At(2, 1)) * sqdip1);
                q.At(0, 2, (DCM.At(0, 1) + DCM.At(1, 0)) * sqdip1);
                q.At(0, 3, (DCM.At(2, 1) + DCM.At(1, 2)) * sqdip1);
            }
        }
        return q;
    }

    public static double trace (Matrix A)
    {
        double tr = (double)A.At(0, 0) + A.At(1, 1) + A.At(2, 2);
        return tr;
    }

    public static Tuple<Matrix, Matrix> KF_Cholesky_update(Matrix P, Matrix v, Matrix R, Matrix H)
    {        
        Matrix PHt = Matrix_Mult(P, Matrix_Transpose(H));
        Matrix S = Matrix_Plus(Matrix_Mult(H, PHt), R);
        S = Matrix_Const_Mult(Matrix_Plus(S, Matrix_Transpose(S)), (double)0.5);
        var SChol = S.Cholesky();
        var SCholInv = SChol.Factor.Inverse();
        Matrix W1 = Matrix_Mult(PHt, (Matrix) SCholInv);
        Matrix W = Matrix_Mult(W1,Matrix_Transpose((Matrix) SCholInv));
        Matrix x = Matrix_Mult(W,v);
        P = Matrix_Minus(P, Matrix_Mult(W1, Matrix_Transpose(W1)));
        return new Tuple<Matrix, Matrix>(x, P);        
    }

}
