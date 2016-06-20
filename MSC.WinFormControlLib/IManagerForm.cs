namespace MSC.WinFormControlLib
{
    using System;

    public interface IManagerForm
    {
        void BindData(int pAutoID, bool pStartEdit);

        bool _IsCanClose { get; set; }
    }
}

