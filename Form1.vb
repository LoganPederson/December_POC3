Imports System.IO
Imports System.Linq
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports NAudio.CoreAudioApi
Imports NAudio.Wave

Public Class Form1
    Private selectedMicDevice As MMDevice
    Private selectedRenderDevice As MMDevice

    Private micCapture As WasapiCapture
    Private loopbackCapture As WasapiLoopbackCapture

    Private micBuffer As New List(Of Byte)
    Private loopbackBuffer As New List(Of Byte)
    Private audioLock As New Object()

    Private stopRecording As Boolean = False
    Private isRecording As Boolean = False
    Private chunkDurationSeconds As Integer = 10
    Private currentChunkId As Integer = 1
    Private chunkStartTime As DateTime

    Private transcriptionResults As New Dictionary(Of Integer, String)
    Private nextToDisplay As Integer = 1

    Private httpClient As HttpClient
    Private uiContext As SynchronizationContext

    Private micDeviceList As New List(Of (Device As MMDevice, Name As String, SR As Integer, Ch As Integer))
    Private loopbackDeviceList As New List(Of (Device As MMDevice, Name As String, SR As Integer, Ch As Integer))
    Private combinedDeviceList As New List(Of (MicDev As (Device As MMDevice, Name As String, SR As Integer, Ch As Integer), LoopDev As (Device As MMDevice, Name As String, SR As Integer, Ch As Integer), DisplayName As String))

    Private finalSampleRate As Integer
    Private finalChannels As Integer

    Private micFormat As WaveFormat
    Private loopFormat As WaveFormat

    ' Unique session identifier for filenames
    Private currentSessionId As String

    ' Track how many samples have been written for mic and loopback to maintain time alignment
    Private lastMicSamples As Integer = 0
    Private lastLoopSamples As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        uiContext = SynchronizationContext.Current
        httpClient = New HttpClient()

        PopulateAudioDevices()
        If cboMicDevice.Items.Count > 0 Then cboMicDevice.SelectedIndex = 0
        If cboLoopbackDevice.Items.Count > 0 Then cboLoopbackDevice.SelectedIndex = 0
        If cboCombinedDevice.Items.Count > 0 Then cboCombinedDevice.SelectedIndex = 0
    End Sub

    Private Sub PopulateAudioDevices()
        Dim enumerator = New MMDeviceEnumerator()

        ' Enumerate capture (mic) devices
        Dim micDevices = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
        For Each dev In micDevices
            Dim f = dev.AudioClient.MixFormat
            micDeviceList.Add((dev, dev.FriendlyName, f.SampleRate, f.Channels))
        Next

        ' Enumerate render (loopback) devices
        Dim renderDevices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
        For Each dev In renderDevices
            Dim f = dev.AudioClient.MixFormat
            loopbackDeviceList.Add((dev, dev.FriendlyName, f.SampleRate, f.Channels))
        Next

        ' Populate mic dropdown
        For Each m In micDeviceList
            cboMicDevice.Items.Add($"{m.Name} (Ch: {m.Ch}, SR: {m.SR} Hz)")
        Next

        ' Populate loopback dropdown: add (LOOPBACK)
        For Each l In loopbackDeviceList
            cboLoopbackDevice.Items.Add($"{l.Name} (Ch: {l.Ch}, SR: {l.SR} Hz) (LOOPBACK)")
        Next

        ' Build combined devices using parentheses matching
        For Each m In micDeviceList
            Dim micParen = ExtractParenthesis(m.Name)
            For Each l In loopbackDeviceList
                Dim loopParen = ExtractParenthesis(l.Name)
                If Not String.IsNullOrEmpty(micParen) AndAlso micParen = loopParen AndAlso m.SR = l.SR Then
                    Dim tempFinalChannels = Math.Max(m.Ch, l.Ch)
                    Dim tempFinalSR = m.SR
                    Dim display = $"{m.Name} (Ch: {m.Ch}, SR: {m.SR} Hz) + {l.Name} (Ch: {l.Ch}, SR: {l.SR} Hz) (LOOPBACK) - Combined (Ch: {tempFinalChannels}, SR: {tempFinalSR} Hz)"
                    combinedDeviceList.Add((m, l, display))
                End If
            Next
        Next

        For Each c In combinedDeviceList
            cboCombinedDevice.Items.Add(c.DisplayName)
        Next
    End Sub

    Private Function ExtractParenthesis(str As String) As String
        Dim start = str.IndexOf("("c)
        Dim [end] = str.IndexOf(")"c)
        If start >= 0 AndAlso [end] > start Then
            Return str.Substring(start + 1, [end] - start - 1).Trim()
        End If
        Return String.Empty
    End Function

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        If isRecording Then
            MessageBox.Show("Already recording")
            Return
        End If

        If cboCombinedDevice.SelectedIndex < 0 Then
            MessageBox.Show("Please select a combined device pairing.")
            Return
        End If

        Dim selectedCombined = combinedDeviceList(cboCombinedDevice.SelectedIndex)
        selectedMicDevice = selectedCombined.MicDev.Device
        selectedRenderDevice = selectedCombined.LoopDev.Device

        finalSampleRate = selectedCombined.MicDev.SR
        finalChannels = Math.Max(selectedCombined.MicDev.Ch, selectedCombined.LoopDev.Ch)

        txtTranscription.Clear()
        txtSummary.Clear()
        transcriptionResults.Clear()
        nextToDisplay = 1
        currentChunkId = 1

        micBuffer.Clear()
        loopbackBuffer.Clear()
        lastMicSamples = 0
        lastLoopSamples = 0

        stopRecording = False
        currentSessionId = DateTime.Now.ToString("yyyyMMdd_HHmmss")

        StartRecording()
    End Sub

    Private Sub StartRecording()
        micCapture = New WasapiCapture(selectedMicDevice)
        AddHandler micCapture.DataAvailable, AddressOf OnMicDataAvailable
        micCapture.StartRecording()
        micFormat = micCapture.WaveFormat

        loopbackCapture = New WasapiLoopbackCapture(selectedRenderDevice)
        AddHandler loopbackCapture.DataAvailable, AddressOf OnLoopbackDataAvailable
        loopbackCapture.StartRecording()
        loopFormat = loopbackCapture.WaveFormat

        isRecording = True
        chunkStartTime = DateTime.Now
    End Sub

    Private Sub OnMicDataAvailable(sender As Object, e As WaveInEventArgs)
        If stopRecording Then Return
        SyncLock audioLock
            AppendDeviceData(e.Buffer, e.BytesRecorded, micFormat, micBuffer, lastMicSamples)
        End SyncLock
        CheckChunkInterval()
    End Sub

    Private Sub OnLoopbackDataAvailable(sender As Object, e As WaveInEventArgs)
        If stopRecording Then Return
        SyncLock audioLock
            AppendDeviceData(e.Buffer, e.BytesRecorded, loopFormat, loopbackBuffer, lastLoopSamples)
        End SyncLock
        CheckChunkInterval()
    End Sub

    Private Sub AppendDeviceData(rawData As Byte(), bytesRecorded As Integer, devFormat As WaveFormat, deviceBuffer As List(Of Byte), ByRef lastSamples As Integer)
        ' Calculate current elapsed samples
        Dim currentElapsedSamples = CInt((DateTime.Now - chunkStartTime).TotalSeconds * finalSampleRate * finalChannels)

        ' If we have a gap, insert silence
        If lastSamples < currentElapsedSamples Then
            Dim silenceNeeded = currentElapsedSamples - lastSamples
            Dim silenceBytes = silenceNeeded * 2 ' each sample 2 bytes
            deviceBuffer.AddRange(New Byte(silenceBytes - 1) {}) ' add silence
            lastSamples = currentElapsedSamples
        End If

        ' Convert incoming data to final PCM format
        Dim converted = ConvertTo16BitPcm(rawData.Take(bytesRecorded).ToArray(), devFormat)
        converted = EnsureChannels(converted, devFormat.Channels, finalChannels)

        ' Append converted data
        deviceBuffer.AddRange(converted)
        ' Update sample count
        lastSamples += (converted.Length / 2)
    End Sub

    Private Sub CheckChunkInterval()
        Dim elapsed = (DateTime.Now - chunkStartTime).TotalSeconds
        If elapsed >= chunkDurationSeconds AndAlso Not stopRecording Then
            SyncLock audioLock
                SaveCurrentChunk(padAtEnd:=True) ' Intermediate chunk, pad to full length
            End SyncLock
        End If
    End Sub

    Private Sub SaveCurrentChunk(padAtEnd As Boolean)
        ' At this point, micBuffer and loopbackBuffer are already time-aligned with silence at start if needed.

        Dim micData = micBuffer.ToArray()
        Dim loopData = loopbackBuffer.ToArray()

        micBuffer.Clear()
        loopbackBuffer.Clear()
        lastMicSamples = 0
        lastLoopSamples = 0

        chunkStartTime = DateTime.Now

        Dim micLen = micData.Length
        Dim loopLen = loopData.Length
        Dim maxLen = Math.Max(micLen, loopLen)

        If padAtEnd Then
            ' For intermediate chunks, pad to full chunk
            Dim samplesNeeded = chunkDurationSeconds * finalSampleRate * finalChannels * 2
            If micLen < samplesNeeded Then
                Dim diff = samplesNeeded - micLen
                Dim temp = New Byte(samplesNeeded - 1) {}
                Buffer.BlockCopy(micData, 0, temp, 0, micLen)
                micData = temp
                micLen = samplesNeeded
            End If
            If loopLen < samplesNeeded Then
                Dim diff = samplesNeeded - loopLen
                Dim temp = New Byte(samplesNeeded - 1) {}
                Buffer.BlockCopy(loopData, 0, temp, 0, loopLen)
                loopData = temp
                loopLen = samplesNeeded
            End If
            maxLen = samplesNeeded
        Else
            ' For final chunk, no padding at the end
            ' Just ensure both arrays are same length by padding silence to the shorter one if needed
            maxLen = Math.Max(micLen, loopLen)
            If micLen < maxLen Then
                Dim temp = New Byte(maxLen - 1) {}
                Buffer.BlockCopy(micData, 0, temp, 0, micLen)
                micData = temp
            End If
            If loopLen < maxLen Then
                Dim temp = New Byte(maxLen - 1) {}
                Buffer.BlockCopy(loopData, 0, temp, 0, loopLen)
                loopData = temp
            End If
        End If

        Dim mixed = MixAudio(micData, loopData, finalChannels)

        Dim filename = $"output_chunk_{currentSessionId}_{currentChunkId}.wav"
        Dim thisChunkId = currentChunkId
        currentChunkId += 1

        Task.Run(Function()
                     SaveAndTranscribeChunk(mixed, finalSampleRate, finalChannels, filename, thisChunkId)
                     Return Task.CompletedTask
                 End Function)
    End Sub

    Private Function ConvertTo16BitPcm(inputData As Byte(), format As WaveFormat) As Byte()
        If inputData.Length = 0 Then
            Return inputData
        End If

        If format.Encoding = WaveFormatEncoding.IeeeFloat Then
            Dim bytesPerSample = format.BitsPerSample / 8
            Dim sampleCount = inputData.Length / bytesPerSample
            Dim outData(sampleCount * 2 - 1) As Byte

            For i As Integer = 0 To sampleCount - 1
                Dim floatVal = BitConverter.ToSingle(inputData, i * 4)
                Dim intVal = CShort(Math.Max(Math.Min(floatVal * Short.MaxValue, Short.MaxValue), Short.MinValue))
                Dim bytes = BitConverter.GetBytes(intVal)
                outData(i * 2) = bytes(0)
                outData(i * 2 + 1) = bytes(1)
            Next
            Return outData
        ElseIf format.Encoding = WaveFormatEncoding.Pcm AndAlso format.BitsPerSample = 16 Then
            Return inputData
        Else
            Return inputData
        End If
    End Function

    Private Function EnsureChannels(data As Byte(), currentChannels As Integer, targetChannels As Integer) As Byte()
        If targetChannels = currentChannels OrElse data.Length = 0 Then
            Return data
        End If

        Dim sampleCount = data.Length / 2
        Dim frameCount = sampleCount / currentChannels
        Dim outData(targetChannels * frameCount * 2 - 1) As Byte

        For f As Integer = 0 To frameCount - 1
            Dim frameSamples As New List(Of Short)
            For c As Integer = 0 To currentChannels - 1
                frameSamples.Add(BitConverter.ToInt16(data, (f * currentChannels + c) * 2))
            Next

            Dim avg = CShort(frameSamples.Average(Function(x) CInt(x)))
            Dim outFrame = Enumerable.Repeat(avg, targetChannels).ToArray()

            For c As Integer = 0 To targetChannels - 1
                Dim bytes = BitConverter.GetBytes(outFrame(c))
                Buffer.BlockCopy(bytes, 0, outData, (f * targetChannels + c) * 2, 2)
            Next
        Next

        Return outData
    End Function

    Private Function MixAudio(micData As Byte(), loopData As Byte(), channels As Integer) As Byte()
        Dim maxLen = Math.Max(micData.Length, loopData.Length)
        Dim micExtended = New Byte(maxLen - 1) {}
        Dim loopExtended = New Byte(maxLen - 1) {}
        Array.Copy(micData, micExtended, micData.Length)
        Array.Copy(loopData, loopExtended, loopData.Length)

        For i As Integer = 0 To maxLen - 2 Step 2
            Dim micSample = BitConverter.ToInt16(micExtended, i)
            Dim loopSample = BitConverter.ToInt16(loopExtended, i)
            Dim mixedSample = CShort((CInt(micSample) + CInt(loopSample)) \ 2)
            Dim bytes = BitConverter.GetBytes(mixedSample)
            micExtended(i) = bytes(0)
            micExtended(i + 1) = bytes(1)
        Next

        Return micExtended
    End Function

    Private Async Function SaveAndTranscribeChunk(frames As Byte(), rate As Integer, ch As Integer, filename As String, chunkId As Integer) As Task
        SaveWav(frames, rate, ch, filename)
        Dim text = Await TranscribeFileAsync(filename)

        SyncLock transcriptionResults
            transcriptionResults(chunkId) = text
        End SyncLock

        uiContext.Post(AddressOf DisplayPendingTranscriptions, Nothing)
    End Function

    Private Sub SaveWav(frames As Byte(), rate As Integer, channels As Integer, filename As String)
        Using wav = New WaveFileWriter(filename, New WaveFormat(rate, 16, channels))
            wav.Write(frames, 0, frames.Length)
        End Using
    End Sub

    Private Sub DisplayPendingTranscriptions()
        While transcriptionResults.ContainsKey(nextToDisplay)
            Dim text = transcriptionResults(nextToDisplay)
            txtTranscription.AppendText(text & vbCrLf)
            nextToDisplay += 1
        End While
    End Sub

    Private Async Function TranscribeFileAsync(filename As String) As Task(Of String)
        Using content As New MultipartFormDataContent()
            Dim fileBytes = File.ReadAllBytes(filename)
            Dim fileContent = New ByteArrayContent(fileBytes)
            fileContent.Headers.ContentType = New MediaTypeHeaderValue("application/octet-stream")
            content.Add(fileContent, "audio", Path.GetFileName(filename))

            Dim request = Await httpClient.PostAsync("YOUR_TRANSCRIPTION_ENDPOINT", content)
            request.EnsureSuccessStatusCode()
            Dim responseStr = Await request.Content.ReadAsStringAsync()

            Return "Transcribed text for " & filename ' Placeholder
        End Using
    End Function

    Private Async Function SummarizeTextAsync(text As String) As Task(Of String)
        Dim payload = "{""messages"":[{""role"":""system"",""content"":[{""type"":""text"",""text"":""Please summarize...""}]},{""role"":""user"",""content"":[{""type"":""text"",""text"":""" & text.Replace("""", "\""") & """}]},{""role"":""assistant"",""content"":[{""type"":""text"",""text"":""""}]}],""temperature"":0.7,""top_p"":0.95,""max_tokens"":800}"
        Dim content = New StringContent(payload, Encoding.UTF8, "application/json")
        Dim resp = Await httpClient.PostAsync("YOUR_SUMMARIZE_ENDPOINT", content)
        resp.EnsureSuccessStatusCode()
        Dim responseStr = Await resp.Content.ReadAsStringAsync()

        Return "Summary of provided transcription" ' Placeholder
    End Function

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        If Not isRecording Then
            MessageBox.Show("Not currently recording")
            Return
        End If

        stopRecording = True

        If micCapture IsNot Nothing Then
            micCapture.StopRecording()
            micCapture.Dispose()
            micCapture = Nothing
        End If

        If loopbackCapture IsNot Nothing Then
            loopbackCapture.StopRecording()
            loopbackCapture.Dispose()
            loopbackCapture = Nothing
        End If

        ' After stopping, if there's leftover data, save final chunk without padding at end
        SyncLock audioLock
            If micBuffer.Count > 0 OrElse loopbackBuffer.Count > 0 Then
                SaveCurrentChunk(padAtEnd:=False)
            End If
        End SyncLock

        isRecording = False

        Dim fullText = txtTranscription.Text
        Task.Run(Async Function()
                     Dim summary = Await SummarizeTextAsync(fullText)
                     uiContext.Post(Sub()
                                        txtSummary.Text = summary
                                    End Sub, Nothing)
                 End Function)
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtTranscription.Clear()
        txtSummary.Clear()
    End Sub
End Class