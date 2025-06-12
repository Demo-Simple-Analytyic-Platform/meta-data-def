DECLARE @nm_secret NVARCHAR(128) = 'nm_secret',
        @ds_secret NVARCHAR(999) = 'ds_secret';

DELETE FROM dbo.secrets WHERE nm_secret = @nm_secret;
INSERT INTO dbo.secrets (nm_secret, ds_secret) VALUES (@nm_secret, @ds_secret);
GO