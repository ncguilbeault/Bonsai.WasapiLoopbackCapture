using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using NAudio;

namespace Bonsai.WasapiLoopbackCapture
{
    public class PlaybackDeviceNameConverter : TypeConverter
    {
        public override bool GetStandardValuesSupported(ITypeDescriptorContext context)
        {
            return true;
        }

        public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext context)
        {
            var deviceEnum = new NAudio.CoreAudioApi.MMDeviceEnumerator();
            var devices = deviceEnum.EnumerateAudioEndPoints(NAudio.CoreAudioApi.DataFlow.All, NAudio.CoreAudioApi.DeviceState.Active).ToList();
            return new StandardValuesCollection(devices);
        }
    }
}