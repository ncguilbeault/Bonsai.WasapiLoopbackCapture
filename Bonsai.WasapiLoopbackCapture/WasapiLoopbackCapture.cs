using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Xml.Serialization;
using System.Globalization;
using Bonsai.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using OpenTK.Audio.OpenAL;
using System.ComponentModel;
using System.Threading;
using OpenTK;
using System.Reactive.Linq;
using System.Reactive.Concurrency;
using System.Reactive.Disposables;
using System.Reactive;
using System.Runtime.InteropServices;
using OpenCV.Net;
using System.Threading.Tasks;
using System.IO;
using NAudio;
using System.Drawing.Design;

namespace Bonsai.WasapiLoopbackCapture
{
    [Description("Produces a sequence of buffered samples acquired from the specified audio capture device.")]
    public class WasapiLoopbackCapture : Source<Mat>
    {

        private string deviceName;
        [Description("The name of the capture device from which to acquire samples.")]
        [TypeConverter(typeof(PlaybackDeviceNameConverter))]
        public string DeviceName { get => deviceName; set => deviceName = value; }

        private int sampleRate; 
        [Description("The sample rate used by the audio capture device, in Hz.")]
        public int SampleRate { get => sampleRate; set => sampleRate = value; }

        private ALFormat sampleFormat;
        [TypeConverter(typeof(SampleFormatConverter))]
        [Description("The requested capture buffer format.")]
        public ALFormat SampleFormat { get => sampleFormat; set => sampleFormat = value; }

        private Depth matDepth;
        [TypeConverter(typeof(MatDepthFormatConverter))]
        [Description("The requested capture buffer format.")]
        public Depth MatDepthFormat { get => matDepth; set => matDepth = value; }

        private int channelCount;
        [Description("Number of channels in the output matrix.")]
        public int Channels { get => channelCount; set => channelCount = value; }

        private double bufferLength;
        private int bufferSize;
        [Description("The length of the capture buffer (ms).")]
        public double BufferLength { get => bufferLength; set { bufferLength = value; bufferSize = (int) Math.Ceiling(sampleRate* value / 1000); } }

        private string fileName;
        [Description("The name of the audio output file.")]
        [FileNameFilter("Audio|*.wav|All Files|*.*")]
        [Editor("Bonsai.Design.SaveFileNameEditor, Bonsai.Design", typeof(UITypeEditor))]
        public string FileName { get => fileName; set => fileName = value; }

        IObservable<Mat> source;
        readonly object captureLock = new object();

        public WasapiLoopbackCapture()
        {
            SampleFormat = ALFormat.Mono16;
            BufferLength = 10;
            SampleRate = 44100;

            source = Observable.Create<Mat>((observer, cancellationToken) =>
            {
                return Task.Factory.StartNew(() =>
                {
                    Console.WriteLine("Initialized.");

                    lock (captureLock)
                    {
                        var devices = new NAudio.CoreAudioApi.MMDeviceEnumerator();
                        Console.WriteLine("Devices: {0}", devices);
                        Console.WriteLine("Device Name: {0}", DeviceName);
                        string device_ID = "";
                        foreach (NAudio.CoreAudioApi.MMDevice device in devices.EnumerateAudioEndPoints(NAudio.CoreAudioApi.DataFlow.All, NAudio.CoreAudioApi.DeviceState.Active).ToList())
                        {
                            if (device.FriendlyName == DeviceName)
                            {
                                device_ID = device.ID;
                            }
                        }
                        NAudio.CoreAudioApi.MMDevice selected_device = devices.GetDevice(device_ID);
                        Console.WriteLine("Selected device: {0}", selected_device);
                        NAudio.CoreAudioApi.WasapiCapture capture = new NAudio.Wave.WasapiLoopbackCapture(selected_device);
                        Console.WriteLine("Creating event handler.");
                        NAudio.Wave.WaveFileWriter writer = new NAudio.Wave.WaveFileWriter(fileName, capture.WaveFormat);
                        capture.DataAvailable += (s, a) =>
                        {
                            Console.WriteLine("Data has become available.");
                            Console.WriteLine("Writing new data to file.");
                            writer.Write(a.Buffer, 0, a.BytesRecorded);
                            unsafe
                            {
                                fixed (byte* p = a.Buffer)
                                {
                                    Console.WriteLine("Audio data copying to buffer.");
                                    Console.WriteLine("Channels: {0}. BitsPerSample: {1}. SampleRate: {2}.", capture.WaveFormat.Channels, capture.WaveFormat.BitsPerSample, capture.WaveFormat.SampleRate);
                                    var buffer = new Mat(channelCount, bufferSize, matDepth, 1);
                                    Console.WriteLine("Matrix Channels: {0}. Matrix Size: {1}.", buffer.Channels, buffer.Size);
                                    Console.WriteLine("Capture Buffer Length: {0}. Capture Bytes Recorded: {1}", a.Buffer.Length, a.BytesRecorded);
                                    buffer.SetData((IntPtr)p, buffer.Step);
                                    Console.WriteLine("Sending audio data to next observer.");
                                    observer.OnNext(buffer);
                                }
                            }
                        };
                        Console.WriteLine("Starting recording.");
                        capture.StartRecording();
                        Console.WriteLine("Recording started.");
                        try
                        {
                            while (!cancellationToken.IsCancellationRequested)
                            {
                                // Do Nothing.
                            }
                            capture.StopRecording();
                        }
                        finally
                        {
                            capture.Dispose();
                            devices.Dispose();
                            selected_device.Dispose();
                            writer.Dispose();
                            writer = null;
                        }
                    }
                },
                cancellationToken,
                TaskCreationOptions.LongRunning,
                TaskScheduler.Default);
            })
            .PublishReconnectable()
            .RefCount();
        }

        public override IObservable<Mat> Generate()
        {
            return source;
        }

        class SampleFormatConverter : EnumConverter
        {
            public SampleFormatConverter(Type type)
                : base(type)
            {
            }

            public override bool GetStandardValuesSupported(ITypeDescriptorContext context)
            {
                return true;
            }

            public override bool GetStandardValuesExclusive(ITypeDescriptorContext context)
            {
                return true;
            }

            public override TypeConverter.StandardValuesCollection GetStandardValues(ITypeDescriptorContext context)
            {
                return new StandardValuesCollection((ALFormat[])Enum.GetValues(typeof(ALFormat)));
            }
        }
        class MatDepthFormatConverter : EnumConverter
        {
            public MatDepthFormatConverter(Type type)
                : base(type)
            {
            }

            public override bool GetStandardValuesSupported(ITypeDescriptorContext context)
            {
                return true;
            }

            public override bool GetStandardValuesExclusive(ITypeDescriptorContext context)
            {
                return true;
            }

            public override TypeConverter.StandardValuesCollection GetStandardValues(ITypeDescriptorContext context)
            {
                return new StandardValuesCollection((Depth[])Enum.GetValues(typeof(Depth)));
            }
        }
    }
}
