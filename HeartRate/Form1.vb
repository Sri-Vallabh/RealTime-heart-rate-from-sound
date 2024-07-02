Imports NAudio.Wave
Imports System.Numerics
Imports MathNet.Numerics.IntegralTransforms
Imports System.Drawing
Imports System.Linq

Public Class Form1
    Private waveIn As WaveInEvent
    Private recordedAudio As List(Of Single)
    Private waveformBitmap As Bitmap
    Private waveformGraphics As Graphics

    ' UI elements (Add these in the Form Designer)
    Private btnStart As Button
    Private btnStop As Button
    Private lblStatus As Label
    Private waveformPictureBox As PictureBox

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize UI elements
        btnStart = New Button With {.Text = "Start Recording", .Location = New Point(10, 10)}
        btnStop = New Button With {.Text = "Stop Recording", .Location = New Point(150, 10)}
        lblStatus = New Label With {.Text = "Status: Not Recording", .Location = New Point(10, 50), .AutoSize = False, .Size = New Size(500, 40)}
        waveformPictureBox = New PictureBox With {.Location = New Point(10, 100), .Size = New Size(500, 200)}

        AddHandler btnStart.Click, AddressOf StartRecording
        AddHandler btnStop.Click, AddressOf StopRecording

        Me.Controls.Add(btnStart)
        Me.Controls.Add(btnStop)
        Me.Controls.Add(lblStatus)
        Me.Controls.Add(waveformPictureBox)

        waveformBitmap = New Bitmap(waveformPictureBox.Width, waveformPictureBox.Height)
        waveformGraphics = Graphics.FromImage(waveformBitmap)
        waveformPictureBox.Image = waveformBitmap
    End Sub

    Private Sub StartRecording(sender As Object, e As EventArgs)
        waveIn = New WaveInEvent()
        waveIn.WaveFormat = New WaveFormat(44100, 1)

        recordedAudio = New List(Of Single)()

        AddHandler waveIn.DataAvailable, AddressOf waveIn_DataAvailable
        AddHandler waveIn.RecordingStopped, AddressOf waveIn_RecordingStopped

        waveIn.StartRecording()
        lblStatus.Text = "Status: Recording..." & Environment.NewLine & "Estimated Heart Rate: N/A"
    End Sub

    Private Sub StopRecording(sender As Object, e As EventArgs)
        waveIn.StopRecording()
        lblStatus.Text = "Status: Processing..." & Environment.NewLine & "Estimated Heart Rate: N/A"
    End Sub

    Private Sub waveIn_DataAvailable(sender As Object, e As WaveInEventArgs)
        For i As Integer = 0 To e.BytesRecorded - 1 Step 2
            recordedAudio.Add(BitConverter.ToInt16(e.Buffer, i) / 32768.0F)
        Next

        ' Update waveform display
        DrawWaveform(recordedAudio)
    End Sub

    Private Sub waveIn_RecordingStopped(sender As Object, e As StoppedEventArgs)
        ProcessAudio(recordedAudio.ToArray())
    End Sub

    Private Sub ProcessAudio(audioData As Single())
        ' Perform STFT and heart rate estimation
        Dim heartRate As Double = EstimateHeartRate(audioData)
        lblStatus.Text = "Status: Not Recording" & Environment.NewLine & "Estimated Heart Rate: " & heartRate & " bpm"
    End Sub

    Private Function EstimateHeartRate(audioData As Single()) As Double
        ' Perform STFT
        Dim sampleRate As Integer = 44100
        Dim nfft As Integer = 2048
        Dim hopSize As Integer = 512
        Dim stft As Complex()() = ComputeSTFT(audioData, nfft, hopSize, sampleRate)

        ' Calculate the heart rate from the frequency domain
        Dim heartRate As Double = CalculateHeartRateFromSTFT(stft, sampleRate, nfft)
        Return heartRate
    End Function

    Private Function ComputeSTFT(audioData As Single(), nfft As Integer, hopSize As Integer, sampleRate As Integer) As Complex()()
        Dim numFrames As Integer = Math.Ceiling(audioData.Length / hopSize)
        Dim stft(numFrames - 1)() As Complex

        For frameIndex As Integer = 0 To numFrames - 1
            Dim frameStart As Integer = frameIndex * hopSize
            Dim frameEnd As Integer = Math.Min(frameStart + nfft, audioData.Length)
            Dim windowedFrame(nfft - 1) As Complex

            For i As Integer = 0 To nfft - 1
                If frameStart + i < frameEnd Then
                    windowedFrame(i) = New Complex(audioData(frameStart + i) * HammingWindow(i, nfft), 0)
                Else
                    windowedFrame(i) = Complex.Zero
                End If
            Next

            stft(frameIndex) = FFT(windowedFrame)
        Next

        Return stft
    End Function

    Private Function HammingWindow(index As Integer, nfft As Integer) As Double
        Return 0.54 - 0.46 * Math.Cos(2 * Math.PI * index / (nfft - 1))
    End Function

    Private Function FFT(windowedFrame As Complex()) As Complex()
        Fourier.Forward(windowedFrame, FourierOptions.NoScaling)
        Return windowedFrame
    End Function

    Private Function CalculateHeartRateFromSTFT(stft As Complex()(), sampleRate As Integer, nfft As Integer) As Double
        Dim minFreq As Double = 0.67
        Dim maxFreq As Double = 2.0

        Dim freqResolution As Double = sampleRate / CDbl(nfft)
        Dim minBin As Integer = CInt(Math.Floor(minFreq / freqResolution))
        Dim maxBin As Integer = CInt(Math.Ceiling(maxFreq / freqResolution))

        Dim hrEnergy As Double() = New Double(maxBin - minBin) {}

        For Each frame As Complex() In stft
            For bin As Integer = minBin To maxBin
                hrEnergy(bin - minBin) += frame(bin).Magnitude
            Next
        Next

        If hrEnergy.Max() = 0 Then
            Return 0 ' Or handle appropriately if no heart rate is detected
        End If

        Dim maxEnergyIdx As Integer = Array.IndexOf(hrEnergy, hrEnergy.Max())
        Dim heartRateFreq As Double = (minBin + maxEnergyIdx) * freqResolution
        Dim heartRateBPM As Double = heartRateFreq * 60

        ' Debug print
        Console.WriteLine("Heart Rate Frequency: " & heartRateFreq & " Hz")
        Console.WriteLine("Heart Rate BPM: " & heartRateBPM)

        Return heartRateBPM
    End Function

    Private Sub DrawWaveform(audioData As List(Of Single))
        waveformGraphics.Clear(Color.Black)
        Dim pen As New Pen(Color.Green)

        Dim midY As Integer = waveformPictureBox.Height / 2
        Dim scaleX As Double = waveformPictureBox.Width / CDbl(audioData.Count)
        Dim scaleY As Double = midY

        For i As Integer = 1 To audioData.Count - 1
            Dim x1 As Integer = CInt((i - 1) * scaleX)
            Dim y1 As Integer = CInt(midY - audioData(i - 1) * scaleY)
            Dim x2 As Integer = CInt(i * scaleX)
            Dim y2 As Integer = CInt(midY - audioData(i) * scaleY)

            waveformGraphics.DrawLine(pen, x1, y1, x2, y2)
        Next

        waveformPictureBox.Invalidate()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If waveIn IsNot Nothing Then
            waveIn.Dispose()
        End If
    End Sub
End Class
