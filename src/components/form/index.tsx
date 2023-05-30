import React, {useState, useRef, useEffect} from 'react';
import IMask from 'imask';
import * as XLSX from 'xlsx';
import {saveAs}  from 'file-saver';
import './style.module.css';


export const Form = () => {

    const [name, setName] = useState<string>('');
    const [email, setEmail] = useState<string>('');
    const [celular, setCelular] = useState<string>('');
    const [cpf, setCpf] = useState<string>('');
    const [selectedFile, setSelectedFile] = useState<File | null>(null)

    const inputFileRef = useRef<HTMLInputElement>(null);
    const celularRef = useRef<HTMLInputElement>(null);
    const cpfRef = useRef<HTMLInputElement>(null);

    const handleName = (e: React.ChangeEvent<HTMLInputElement>) => {
      setName(e.target.value);
    }

    const handleEmail = (e: React.ChangeEvent<HTMLInputElement>) => {
      setEmail(e.target.value);
    }
    const handleCelular = (e: React.ChangeEvent<HTMLInputElement>) => {
      setCelular(e.target.value);
    }
    const handleCpf = (e: React.ChangeEvent<HTMLInputElement>) => {
      setCpf(e.target.value);
    }
    
    useEffect(()=>{
      const celularMask = IMask(celularRef.current!, {
        mask: '(00) 00000-0000'
      });
      const cpfMask = IMask(cpfRef.current!, {
        mask: '000.000.000-00'
      });
  
      celularMask.on('accept', () => {
        setCelular(celularMask.value);
      });
      cpfMask.on('accept', () => {
        setCpf(cpfMask.value);
      });
  
    }, []);

    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.currentTarget.files?.[0];

      if(file) {
        setSelectedFile(file);
      }
    };
  
    const handleSubmit = (e: React.FormEvent<HTMLFormElement>): void => {
      e.preventDefault();
  
      const formData = {
        name,
        email,
        celular,
        cpf
      }
      
      if(
        name.length > 0
        && email.length > 0
        && celular.length > 0
        && cpf.length > 0
        ) {
          saveSpreadSheet(formData);

          setName('');
          setEmail('');
          setCelular('');
          setCpf('');
      } else {
        alert('Preencha todos os campos');
      }
    }

    const handleSubmitFile = (e: React.MouseEvent<HTMLButtonElement, MouseEvent>): void => {
      e.preventDefault();
  
      if (!selectedFile) {
        alert('Selecione um arquivo para fazer o append dos dados.');
        return;
      }
  
      const formData = {
        name,
        email,
        celular,
        cpf,
      };
  
      appendToSpreadSheet(selectedFile, formData);
    };

    const saveSpreadSheet = (data: any) => {
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.json_to_sheet([data]);
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Dados do Formulário');
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const excelData = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(excelData, 'formulario.xlsx');
    };

    const appendToSpreadSheet = (file: File, data: any) => {
      const reader = new FileReader();
    
      reader.onload = (e) => {
        const workbook = XLSX.read(e.target?.result, { type: 'binary' });
    
        const worksheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[worksheetName];
    
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const headers = jsonData[0] as string[];
    
        const newRow = headers.map((header) => data[header] || '');
    
        const rowIndex = jsonData.length;
        XLSX.utils.sheet_add_aoa(worksheet, [newRow], {
          origin: -1,
        });
    
        const updatedWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(updatedWorkbook, worksheet, worksheetName);
    
        const excelBuffer = XLSX.write(updatedWorkbook, {
          bookType: 'xlsx',
          type: 'array',
        });
        const excelData = new Blob([excelBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
    
        saveAs(excelData, file.name);
      };
    
      reader.readAsBinaryString(file);
    };
    
    
    
    return (
      <div>
        <h1>Formulário</h1>
        <form onSubmit={handleSubmit}>
          <fieldset>
            <label htmlFor="name">
                Nome <br />
                <input type="text" name='name' value={name} onChange={handleName} autoFocus placeholder='Digite seu nome' />
            </label>
            <br />
            <label htmlFor="email">
              Email <br />
              <input type="email" name='email' value={email} onChange={handleEmail} placeholder='Digite seu email' />
            </label>
            <br />
            <label htmlFor="celular">
              Celular <br />
              <input type="text" name='celular' value={celular} ref={celularRef} onChange={handleCelular} placeholder='(00) 00000-0000' />
            </label>
            <br />
            <label htmlFor="cpf">
              CPF <br />
              <input type="text" name='cpf' value={cpf} ref={cpfRef} onChange={handleCpf} placeholder='000.000.000-00'/>
            </label>
            <br />
            <label htmlFor="selectFile">
              Selecione o arquivo <br />
              <input type="file" ref={inputFileRef} onChange={handleFileChange} accept='.xlsx'/>
            </label>
            
            <input type="submit" value={'Salvar'} />
            <button type="button"  onClick={handleSubmitFile}>Incrementar</button>
            
          </fieldset>
        </form>
        
      </div>
    )
  }