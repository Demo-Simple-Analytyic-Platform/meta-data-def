BEGIN
  INSERT INTO tsa_srd.tsa_dq_result_status (id_model, id_dq_result_status, fn_dq_result_status, fd_dq_result_status) VALUES ('<id_model>', '02060d000f060a05030d010006190f03', 'NOK', 'Not Oke, Data Item(s) were measured and DIT NOT passed the defined test.');
  INSERT INTO tsa_srd.tsa_dq_result_status (id_model, id_dq_result_status, fn_dq_result_status, fd_dq_result_status) VALUES ('<id_model>', '060400020703080002000b0406021506', 'OOS', 'Record is "Out-of_scope".');
  INSERT INTO tsa_srd.tsa_dq_result_status (id_model, id_dq_result_status, fn_dq_result_status, fd_dq_result_status) VALUES ('<id_model>', '0e01000005070c020e000d0901190a05', 'OKE', 'Oke, Data Item(s) were measured and passed the defined test.');
END
GO

