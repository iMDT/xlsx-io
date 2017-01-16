package br.com.imdt.xlsx.io;

/**
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public abstract class DefaultCallback implements DataCallback{

    public DefaultCallback() {
    }

    @Override
    public void onBegin() {}

    @Override
    public void onEnd() {}

    @Override
    public void onSheetBegin() {}

    @Override
    public void onSheetEnd() {}
}
