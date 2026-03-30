using System.Collections.Generic;
using SmartSAP.ViewModels;

namespace SmartSAP.ViewModels.Modules
{
    public enum ParameterType
    {
        Text,
        Boolean,
        Choice
    }

    public class StepParameter : ViewModelBase
    {
        private string _name = string.Empty;
        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        private ParameterType _type = ParameterType.Text;
        public ParameterType Type
        {
            get => _type;
            set => SetProperty(ref _type, value);
        }

        private object? _value;
        public object? Value
        {
            get => _value;
            set => SetProperty(ref _value, value);
        }

        private IEnumerable<string>? _options;
        public IEnumerable<string>? Options
        {
            get => _options;
            set => SetProperty(ref _options, value);
        }

        public StepParameter(string name, ParameterType type, object? defaultValue = null, IEnumerable<string>? options = null)
        {
            Name = name;
            Type = type;
            Value = defaultValue;
            Options = options;
        }
    }
}
