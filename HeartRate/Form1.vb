Imports NAudio.Wave
Imports System.Numerics
Imports MathNet.Numerics.IntegralTransforms
Imports System.Drawing
Imports System.Linq
Imports OxyPlot
Imports OxyPlot.Axes
Imports OxyPlot.Series
Imports OxyPlot.WindowsForms
Imports System.IO

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
    Private plotView As PlotView

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized

        ' Initialize UI elements
        btnStart = New Button With {.Text = "Start Recording", .Location = New Point(10, 10)}
        btnStop = New Button With {.Text = "Stop Recording", .Location = New Point(150, 10)}
        lblStatus = New Label With {.Text = "Status: Not Recording", .Location = New Point(10, 50), .AutoSize = False, .Size = New Size(500, 40)}
        waveformPictureBox = New PictureBox With {.Location = New Point(10, 100), .Size = New Size(500, 200)}
        plotView = New PlotView With {.Location = New Point(700, 10), .Size = New Size(1200, 1000)}

        AddHandler btnStart.Click, AddressOf StartRecording
        AddHandler btnStop.Click, AddressOf StopRecording

        Me.Controls.Add(btnStart)
        Me.Controls.Add(btnStop)
        Me.Controls.Add(lblStatus)
        Me.Controls.Add(waveformPictureBox)
        Me.Controls.Add(plotView)

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
        ' Display the STFT and estimate heart rate
        DisplaySTFT(audioData)
    End Sub

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

    Private Sub DisplaySTFT(audioData As Single())
        ' Perform STFT
        Dim sampleRate As Integer = 16000
        Dim n_fft As Integer = 256
        Dim hopSize As Integer = 64
        Dim stftResults As Complex()() = ComputeSTFT(audioData, n_fft, hopSize, sampleRate)

        ' Compute magnitudes and scale from 0 to 255
        Dim maxMagnitude As Double = 0
        For Each frame In stftResults
            For Each c In frame
                maxMagnitude = Math.Max(maxMagnitude, c.Magnitude())
            Next
        Next

        Dim magnitudes(stftResults.Count - 1)() As Double
        For i As Integer = 0 To stftResults.Count - 1
            magnitudes(i) = New Double(stftResults(i).Length - 1) {}
            For j As Integer = 0 To stftResults(i).Length - 1
                magnitudes(i)(j) = 255 * (stftResults(i)(j).Magnitude() / maxMagnitude)
            Next
        Next

        ' Determine the range to display
        Dim displayLength As Integer = n_fft / 2

        Dim frequencies(displayLength - 1) As Double
        Dim nyquist As Double = sampleRate / 2.0

        ' Calculate the frequency resolution
        Dim freqResolution As Double = sampleRate / n_fft

        ' Calculate total time spanned by all frames including initial offset
        Dim totalTime As Double = stftResults.Length * hopSize / sampleRate

        ' Plot the spectrogram
        Dim plotModel As New PlotModel With {
            .Title = "STFT Magnitude",
            .Background = OxyColors.Black,
            .TextColor = OxyColors.White,
            .PlotAreaBorderColor = OxyColors.White
        }

        ' Create and configure ColorAxis for greyscale
        Dim colorAxis As New LinearColorAxis With {
            .Position = AxisPosition.Right,
            .Palette = OxyPalettes.Gray(256),
            .LowColor = OxyColors.Black,
            .HighColor = OxyColors.White,
            .Minimum = 0,
            .Maximum = 255
        }
        plotModel.Axes.Add(colorAxis)

        ' Create and configure HeatMapSeries
        Dim heatMapSeries As New HeatMapSeries With {
            .X0 = 0,
            .X1 = totalTime / 2, ' Adjust X-axis span to total time
            .Y0 = 0,
            .Y1 = nyquist * 2, ' Adjust Y-axis span to Nyquist frequency
            .Data = New Double(magnitudes.Length - 1, displayLength - 1) {}
        }

        ' Fill the heatmap data
        For i As Integer = 0 To magnitudes.Length - 1
            For j As Integer = 0 To displayLength - 1
                heatMapSeries.Data(i, j) = magnitudes(i)(j)
            Next
        Next

        plotModel.Series.Add(heatMapSeries)

        ' Configure X and Y axes
        Dim xAxis As New LinearAxis With {
            .Position = AxisPosition.Bottom,
            .Title = "Time (s)",
            .TextColor = OxyColors.White,
            .TitleColor = OxyColors.White,
            .TicklineColor = OxyColors.White
        }
        plotModel.Axes.Add(xAxis)

        Dim yAxis As New LinearAxis With {
            .Position = AxisPosition.Left,
            .Title = "Frequency (Hz)",
            .TextColor = OxyColors.White,
            .TitleColor = OxyColors.White,
            .TicklineColor = OxyColors.White
        }
        plotModel.Axes.Add(yAxis)

        plotView.Model = plotModel

        ' Estimate heart rate from STFT heatmap
        Dim heartRate As Double = EstimateHeartRateFromHeatmap(heatMapSeries.Data, n_fft, displayLength)
        lblStatus.Text = "Status: Not Recording" & Environment.NewLine & "Estimated Heart Rate: " & heartRate & " bpm"
    End Sub

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

    Private Function EstimateHeartRateFromHeatmap(heatmapData As Double(,), nfft As Integer, displayLength As Integer) As Double
        Dim minFreq As Double = 0.67
        Dim maxFreq As Double = 2.0

        ' Find the corresponding bins for minFreq and maxFreq
        Dim minBin As Integer = CInt(Math.Floor(minFreq * displayLength))
        Dim maxBin As Integer = CInt(Math.Floor(maxFreq * displayLength))

        If maxBin >= displayLength Then
            maxBin = displayLength - 1
        End If

        ' Find the bin with the maximum magnitude in the specified frequency range
        Dim maxMagnitude As Double = 0
        Dim maxBinIndex As Integer = minBin

        For i As Integer = 0 To heatmapData.GetLength(0) - 1
            For j As Integer = minBin To maxBin
                If heatmapData(i, j) > maxMagnitude Then
                    maxMagnitude = heatmapData(i, j)
                    maxBinIndex = j
                End If
            Next
        Next

        ' Convert the bin index to heart rate in bpm
        Dim heartRateFrequency As Double = maxBinIndex * 2 / displayLength
        Dim heartRateBPM As Double = heartRateFrequency * 60

        Return heartRateBPM
    End Function

End Class
