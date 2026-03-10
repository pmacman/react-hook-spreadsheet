import useExcelDownload from '@hooks/ExcelDownloadHook';
import type { Address, ContactInfo } from '@models/ContactInfo';
import type { ExcelDownloadColumn, ExcelDownloadWorkbook } from '@models/ExcelDownload';
import ContactInfoService from '@services/ContactInfoService';
import { getFormattedDate } from '@utils/dateFormatter';
import { useSnackbar } from 'notistack';
import { useCallback, useState } from 'react';

const contactInfoColumns: Array<ExcelDownloadColumn<ContactInfo>> = [
  {
    key: 'firstName',
    label: 'First Name',
  },
  {
    key: 'lastName',
    label: 'Last Name',
  },
  {
    key: 'age',
    label: 'Age',
    render: (age: number | undefined) => (age !== undefined ? age.toString() : 'N/A'),
  },
  {
    key: 'hireDate',
    label: 'Hire Date',
    render: (hireDate: string) => getFormattedDate(hireDate, 'MM/dd/yyyy'),
  },
  {
    key: 'address',
    altKey: 'street1',
    label: 'Address Line 1',
    render: (address: Address) => address.street1,
  },
  {
    key: 'address',
    altKey: 'street2',
    label: 'Address Line 2',
    width: 100,
    render: (address: Address) => address?.street2 ?? '',
  },
  {
    key: 'address',
    altKey: 'city',
    label: 'City',
    render: (address: Address) => address.city,
  },
  {
    key: 'address',
    altKey: 'state',
    label: 'State',
    render: (address: Address) => address?.state ?? '',
  },
  {
    key: 'address',
    altKey: 'country',
    label: 'Country',
    render: (address: Address) => address.country,
  },
];

function Spreadsheet() {
  const { enqueueSnackbar } = useSnackbar();
  const [fileLoading, setFileLoading] = useState<boolean>(false);
  const [downloadExcel] = useExcelDownload<ContactInfo>();

  const handleDownloadClick = useCallback(() => {
    setFileLoading(true);

    ContactInfoService.getContactInfo()
      .then((contactInfoList: ContactInfo[]) => {
        const workbook: ExcelDownloadWorkbook<ContactInfo> = {
          filename: 'react-hook-spreadsheet',
          worksheets: [
            {
              list: contactInfoList,
              worksheetName: 'Contact Info',
              columns: contactInfoColumns,
              defaultColumnWidth: 'auto',
            },
          ],
        };

        downloadExcel(workbook);

        setFileLoading(false);

        enqueueSnackbar('File downloaded!', {
          variant: 'success',
        });
      })
      .catch((errorMessage: string) => {
        enqueueSnackbar(`An error occurred: ${errorMessage}`, {
          variant: 'error',
        });
        setFileLoading(false);
      });
  }, [downloadExcel, enqueueSnackbar]);

  return (
    <button className={'btn'} onClick={handleDownloadClick} disabled={fileLoading}>
      CLICK TO DOWNLOAD CONTACT INFO
    </button>
  );
}

export default Spreadsheet;
